using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.IO;
using System.Threading;

namespace nInterface
{
    class iniFileHandle
    {
        /// <summary>
        /// Registry 형식의 파일 조작에 필요한 기능을 제공합니다.
        /// </summary>
        public interface IPrivateProfileProvider
        {
            /// <summary>
            /// 섹션 리스트를 가져옵니다.
            /// </summary>
            /// <returns>섹션 문자열 배열입니다.</returns>
            string[] GetSectionNames();
            /// <summary>
            /// 섹션을 이용해 키와 값의 Pair 배열을 가져 옵니다.
            /// </summary>
            /// <param name="szSection">섹션 문자열입니다.</param>
            /// <returns>키와 값의 Pair 배열입니다. 하나의 Pair는 '키=값'의 형식으로 가져오게 됩니다.</returns>
            string[] GetPairsBySection(string szSection);
            /// <summary>
            /// 섹션을 삭제합니다.
            /// </summary>
            /// <param name="szSection"></param>
            void DeleteSection(string szSection);
            /// <summary>
            /// 섹션명 밑에 있는 키를 삭제 합니다.
            /// </summary>
            /// <param name="szSection">섹션명입니다.</param>
            /// <param name="szKey">삭제할 키입니다.</param>
            bool DeleteKey(string szSection, string szKey);

            /// <summary>
            /// 섹션명 밑에 키와 값을 파일에 씁니다.
            /// </summary>
            /// <param name="szSection">섹션명입니다.</param>
            /// <param name="szKey">키입니다.</param>
            /// <param name="szValue">값입니다.</param>
            bool WritePair(string szSection, string szKey, string szValue);
        }
        /// <summary>
        /// 레지스트리 형태의 INI 파일 데이터를 접근 하는 Win32 함수를 C# 계층에서 사용 할 수 있도록 wrapping한 클래스입니다.
        /// </summary>
        internal static class Win32RegNative
        {
            /// <summary>
            /// INI 파일에 섹션과 키로 검색하여 값을 저장합니다.
            /// </summary>
            /// <param name="lpAppname">섹션명입니다.</param>
            /// <param name="lpKeyName">키값명입니다.</param>
            /// <param name="lpString">저장 할 문자열입니다.</param>
            /// <param name="lpFileName">파일 이름입니다.</param>
            /// <returns>처리 여부입니다.</returns>
            [DllImport("kernel32.dll")]
            public static extern bool WritePrivateProfileString(string lpAppName, string lpKeyName, string lpString, string lpFileName);

            /// <summary>
            /// INI 파일에 섹션을 저장합니다.
            /// </summary>
            /// <param name="lpAppname">섹션명입니다.</param>
            /// <param name="lpString">키=값으로 되어 있는 문자열 데이터입니다.</param>
            /// <param name="lpFileName">파일 이름입니다.</param>
            /// <returns>처리 여부입니다.</returns>
            [DllImport("kernel32.dll")]
            public static extern bool WritePrivateProfileSection(string lpAppName, string lpString, string lpFileName);

            /// <summary>
            /// INI 파일에 섹션과 키로 검색하여 값을 Integer형으로 읽어 옵니다.
            /// </summary>
            /// <param name="lpAppname">섹션명입니다.</param>
            /// <param name="lpKeyName">키값명입니다.</param>
            /// <param name="nDefalut">기본값입니다.</param>
            /// <param name="lpFileName">파일 이름입니다.</param>
            /// <returns>입력된 값입니다. 만약 해당 키로 검색 실패시 기본 값으로 대체 됩니다.</returns>
            [DllImport("kernel32.dll")]
            public static extern uint GetPrivateProfileInt(string lpAppName, string lpKeyName, int nDefalut, string lpFileName);

            /// <summary>
            /// INI 파일에 섹션과 키로 검색하여 값을 문자열형으로 읽어 옵니다.
            /// </summary>
            /// <param name="lpAppname">섹션명입니다.</param>
            /// <param name="lpKeyName">키값명입니다.</param>
            /// <param name="lpDefault">기본 문자열입니다.</param>
            /// <param name="lpReturnedString">가져온 문자열입니다.</param>
            /// <param name="nSize">문자열 버퍼의 크기입니다.</param>
            /// <param name="lpFileName">파일 이름입니다.</param>
            /// <returns>가져온 문자열 크기입니다.</returns>
            [DllImport("kernel32.dll")]
            public static extern uint GetPrivateProfileString(string lpAppName, string lpKeyName, string lpDefault, StringBuilder lpReturnedString, uint nSize, string lpFileName);
            /// <summary>
            /// INI 파일에 섹션으로 검색하여 키와 값을 Pair형태로 가져옵니다.
            /// </summary>
            /// <param name="lpAppName">섹션명입니다.</param>
            /// <param name="lpPairVaules">Pair한 키와 값을 담을 배열입니다.</param>
            /// <param name="nSize">배열의 크기입니다.</param>
            /// <param name="lpFileName">파일 이름입니다.</param>
            /// <returns>읽어온 바이트 수 입니다.</returns>
            [DllImport("kernel32.dll")]
            public static extern uint GetPrivateProfileSection(string lpAppName, byte[] lpPairVaules, uint nSize, string lpFileName);

            /// <summary>
            /// INI 파일의 섹션을 가져옵니다.
            /// </summary>
            /// <param name="lpSections">섹션의 리스트를 직렬화하여 담을 배열입니다.</param>
            /// <param name="nSize">배열의 크기입니다.</param>
            /// <param name="lpFileName">파일 이름입니다.</param>
            /// <returns>읽어온 바이트 수 입니다.</returns>
            [DllImport("kernel32.dll")]
            public static extern uint GetPrivateProfileSectionNames(byte[] lpSections, uint nSize, string lpFileName);
        }
        /// <summary>
        /// Registry 형태의 INI 파일을 제어하는 기능을 제공합니다.
        /// </summary>
        public class Win32Reg : IPrivateProfileProvider
        {
            /// <summary>
            /// 섹션을 읽는 버퍼의 크기입니다.
            /// </summary>
            public static readonly uint SectionBufferSize = 1024;
            /// <summary>
            /// 키와 값의 Pair를 읽는 버퍼의 크기입니다. 
            /// </summary>
            public static readonly uint PairBufferSize = 16384;

            protected string _szFileName = null;

            protected Win32Reg() { }
            public Win32Reg(string szFileName)
            {
                _szFileName = szFileName;
            }

            public string szFileName
            {
                get { return _szFileName; }
                set { _szFileName = value; }
            }
            /// <summary>
            /// 섹션 리스트를 가져옵니다.
            /// </summary>
            /// <returns>섹션 문자열 배열입니다.</returns>
            public string[] GetSectionNames()
            {
                if (!System.IO.File.Exists(_szFileName)) throw new FileNotFoundException();

                byte[] bSection = new byte[SectionBufferSize];
                if (Win32RegNative.GetPrivateProfileSectionNames(bSection, SectionBufferSize, _szFileName) <= 0)
                {
                    return null;
                }
                return System.Text.Encoding.Default.GetString(bSection).Split(new char[1] { '\0' }, StringSplitOptions.RemoveEmptyEntries);
            }
            /// <summary>
            /// 섹션을 이용해 키와 값의 Pair 배열을 가져 옵니다.
            /// </summary>
            /// <param name="szSection">섹션 문자열입니다.</param>
            /// <returns>키와 값의 Pair 배열입니다. 하나의 Pair는 '키=값'의 형식으로 가져오게 됩니다.</returns>
            public string[] GetPairsBySection(string szSection)
            {
                if (!System.IO.File.Exists(_szFileName)) throw new FileNotFoundException();
                byte[] bPair = new byte[PairBufferSize];
                if (Win32RegNative.GetPrivateProfileSection(szSection, bPair, PairBufferSize, _szFileName) <= 0)
                {
                    return null;
                }
                return System.Text.Encoding.Default.GetString(bPair).Split(new char[1] { '\0' }, StringSplitOptions.RemoveEmptyEntries);
            }
            /// <summary>
            /// 섹션을 삭제합니다.
            /// </summary>
            /// <param name="szSection">삭제 할 섹션명입니다.</param>
            public void DeleteSection(string szSection)
            {
                Win32RegNative.WritePrivateProfileSection(szSection, null, _szFileName);
            }
            /// <summary>
            /// 섹션명 밑에 있는 키를 삭제 합니다.
            /// </summary>
            /// <param name="szSection">섹션명입니다.</param>
            /// <param name="szKey">삭제할 키입니다.</param>
            /// <returns>처리 여부입니다.</returns>
            public bool DeleteKey(string szSection, string szKey)
            {
                return Win32RegNative.WritePrivateProfileString(szSection, szKey, null, _szFileName);
            }
            /// <summary>
            /// 섹션명 밑에 키와 값을 파일에 씁니다.
            /// </summary>
            /// <param name="szSection">섹션명입니다.</param>
            /// <param name="szKey">키입니다.</param>
            /// <param name="szValue">값입니다.</param>
            /// <returns>처리 여부입니다.</returns>
            public bool WritePair(string szSection, string szKey, string szValue)
            {
                return Win32RegNative.WritePrivateProfileString(szSection, szKey, szValue, _szFileName);
            }
            /// <summary>
            /// '키=값'형태의 문자열을 키와 값으로 분리합니다.
            /// </summary>
            /// <param name="szSource">'키=값'의 형태로 이루어진 소스 문자열입니다.</param>
            /// <param name="szKey">키를 담을 String 변수입니다.</param>
            /// <param name="szValue">값을 담을 String 변수입니다.</param>
            public static string[] SplitPair(string szSource)
            {
                string[] szTemp = szSource.Split(new char[1] { '=' }, StringSplitOptions.RemoveEmptyEntries);

                if (szTemp.Length > 2 || szTemp.Length <= 0)
                    throw new ArgumentOutOfRangeException
                    (szSource,
                     "'키=값'으로 된 문자열 형태가 아닙니다.");
                return szTemp;

            }
            /// <summary>
            /// 키와 값을 '키=값'의 형태인 단일문자열로 저장합니다.
            /// </summary>
            /// <param name="szKey">키입니다.</param>
            /// <param name="szValue">값입니다.</param>
            /// <returns>Pair 문자열입니다.</returns>
            public static string MergePair(string szKey, string szValue)
            {
                return szKey + "=" + szValue;
            }
        }
    }
}
