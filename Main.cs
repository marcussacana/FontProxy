using Microsoft.Win32;
using System;
using System.DirectoryServices.AccountManagement;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Runtime.InteropServices;

namespace VNX.FontProxy {
    public static class FontProxy {
        public static string CustomFontPrefix = "_force";
        private static RegistryKey FontSubstitutes = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\FontSubstitutes", true);
        private static RegistryKey Fonts = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Fonts", true);
        private static RegistryKey reg = Registry.LocalMachine.CreateSubKey("SOFTWARE\\VNX\\FontProxy");
        private static string FontDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Fonts) + "\\";
        /// <summary>
        /// Reboot The Computer on make Any Change
        /// </summary>
        public static bool RebootComputer = true;

        private static void Reboot() {
            if (RebootComputer)
                Windows.Reboot();
        }

        /// <summary>
        /// Redirect a Font
        /// </summary>
        /// <param name="OriginalFamily">Original Font Face Name</param>
        /// <param name="TargetFamily">Target Font Face Name</param>
        /// <param name="Status">Enable or Disable the Redirection</param>
        public static void RedirectFont(string OriginalFamily, string TargetFamily, bool Status) {
            OriginalFamily = ParseVar(OriginalFamily, VarType.FaceName);
            TargetFamily = ParseVar(TargetFamily, VarType.FaceName);
            if (Status)
                FontSubstitutes.SetValue(OriginalFamily, TargetFamily);
            else
                FontSubstitutes.DeleteValue(OriginalFamily);
            Reboot();
        }

        /// <summary>
        /// Replace a font without redirection
        /// </summary>
        /// <param name="Original">Original Font</param>
        /// <param name="New">Target Font</param>
        public static void ReplaceFont(string Original, string New) {
            Original = ParseVar(Original, VarType.FaceName);
            if (Original == null)
                throw new Exception();
            VarType NT = DetectType(New);
            bool TargetExist = File.Exists(FontDirectory + New);
            if (NT == VarType.FontFileName && !TargetExist)
                throw new Exception("Invalid Target");
            else
                if (NT == VarType.UnistalledFontPath) {
                string from = ParseVar(New, TargetExist ? VarType.InstalledFontPath : VarType.UnistalledFontPath);
                New = ParseVar(New, VarType.InstalledFontPath);
                Copy(from, New);
            }
            Fonts.SetValue(Original, ParseVar(New, VarType.FontFileName));
            Reboot();
        }

        private static string ParseVar(string var, VarType target) {
            VarType VT = DetectType(var);
            switch (VT) {
                case VarType.FaceName:
                    switch (target) {
                        case VarType.FaceName:
                            return var;
                        case VarType.InstalledFontPath:
                            return FontDirectory + Fonts.GetValue(var);
                        case VarType.FontFileName:
                            return (string)Fonts.GetValue(var);
                        default:
                            throw new Exception("Invalid Input");
                    }
                case VarType.FontFileName:
                    switch (target) {
                        case VarType.FaceName:
                            return SearchFont(var);
                        case VarType.FontFileName:
                            return var;
                        case VarType.InstalledFontPath:
                            return FontDirectory + var;
                        default:
                            throw new Exception("Invalid Input");
                    }
                case VarType.InstalledFontPath:
                    switch (target) {
                        case VarType.FaceName:
                            return GetFontFace(File.ReadAllBytes(var));
                        case VarType.FontFileName:
                            return Path.GetFileName(var);
                        case VarType.InstalledFontPath:
                            return var;
                        default:
                            throw new Exception("Invalid Input");
                    }
                case VarType.UnistalledFontPath:
                    switch (target) {
                        case VarType.FaceName:
                            return GetFontFace(File.ReadAllBytes(var));
                        case VarType.FontFileName:
                            return Path.GetFileName(var);
                        case VarType.InstalledFontPath:
                            return FontDirectory + Path.GetFileNameWithoutExtension(var) + CustomFontPrefix + Path.GetExtension(var);
                        case VarType.UnistalledFontPath:
                            return var;
                        default:
                            throw new Exception("Invalid Input");
                    }
            }
            throw new Exception("Invalid Input");
        }

        public static void Copy(string from, string to) {
            if (File.Exists(to))
                return;
            if (!File.Exists(from))
                return;
            File.Copy(from, to);
        }

        /// <summary>
        /// Install a Font
        /// </summary>
        /// <param name="Path"></param>
        public static void InstallFont(string Path) {
            if (DetectType(Path) != VarType.UnistalledFontPath)
                return;
            string SysPath = FontDirectory + System.IO.Path.GetFileName(Path);
            if (File.Exists(SysPath))
                return;
            Copy(Path, SysPath);
            string FaceName = GetFontFace(File.ReadAllBytes(Path));
            Fonts.SetValue(FaceName, System.IO.Path.GetFileName(Path));
            Reboot();
        }

        private static string GetFontFace(byte[] binary) {
            PrivateFontCollection fc = new PrivateFontCollection();
            IntPtr pointer = Marshal.UnsafeAddrOfPinnedArrayElement(binary, 0);
            fc.AddMemoryFont(pointer, Convert.ToInt32(binary.Length));
            System.Drawing.Font f = new System.Drawing.Font(fc.Families[0], 10);
            FontFamily ff = f.FontFamily;
            return ff.Name;
        }

        private enum VarType {
            UnistalledFontPath,
            FaceName,
            InstalledFontPath,
            FontFileName
        }
        private static VarType DetectType(string Type) {
            if (Type.Length > 2)
                if (!Type.StartsWith(FontDirectory) && Type[1] == ':')
                    return VarType.UnistalledFontPath;
                else if (Type.StartsWith(FontDirectory))
                    return VarType.InstalledFontPath;
                else if (Type.EndsWith(".otf") || Type.EndsWith(".ttf") || Type.EndsWith(".ttc") & Type[1] != ':')
                    return VarType.FontFileName;
            return VarType.FaceName;
        }

        /// <summary>
        /// Get all installed fonts with specified parameters
        /// </summary>
        /// <param name="Monospaced">Search only Monospaced Fonts</param>
        /// <param name="FontCharset">Charset to Search Fonts</param>
        /// <returns>All fonts with specified Charset</returns>
        public static string[] FindFamilies(bool Monospaced, Charset FontCharset) {
            FontWorker FW = new FontWorker();
            return FW.GetFonts(FontCharset, Monospaced);
        }

        /// <summary>
        /// Search for a font
        /// </summary>
        /// <param name="Font">A part of the Font File Name or Face Name</param>
        /// <returns>Family Name</returns>
        public static string SearchFont(string Font) {
            string ff = string.Empty;
            if (VarType.FontFileName == DetectType(Font)) {
                string tmp = ParseVar(Font, VarType.InstalledFontPath);
                if (File.Exists(tmp))
                    ff = GetFontFace(File.ReadAllBytes(tmp));
                if (Font.Contains("."))
                    Font = Path.GetFileNameWithoutExtension(Font);
            }
            Font = Font.ToLower();
            foreach (string f in Fonts.GetValueNames()) {
                string Family = f.ToLower();
                string Value = ((string)Fonts.GetValue(f)).ToLower();
                if (Family.Contains(Font) || Value.Contains(Font) || (!string.IsNullOrEmpty(ff) && Family.Contains(ff)))
                    return f;
            }
            return null;
        }

        /// <summary>
        /// Get a Status of Font Famliy Name
        /// </summary>
        /// <param name="Family">Font Familiy Name</param>
        /// <param name="Search">Search the Family in the Face Name (if Contains)</param>
        /// <returns>Font Status</returns>
        public static FontStatus GetFontStatus(string Family, bool Search) {
            VarType FT = DetectType(Family);
            switch (FT) {
                case VarType.FontFileName:
                    Family = GetFontFace(File.ReadAllBytes(FontDirectory + Family));
                    break;
                case VarType.InstalledFontPath:
                    Family = GetFontFace(File.ReadAllBytes(Family));
                    break;
                case VarType.FaceName:
                    break;
                default:
                    return FontStatus.Unknown;
            }
            Family = Family.ToLower();
            string[] Redirections = FontSubstitutes.GetValueNames();
            foreach (string r in Redirections) {
                string Original = r.ToLower();
                if (Original == Family)
                    return FontStatus.Redirected;
            }

            string[] Families = Fonts.GetValueNames();
            foreach (string f in Families) {
                string Font = f.ToLower();
                string FName = ((string)Fonts.GetValue(f)).ToLower();
                if (Search && Font.Contains(Family))
                    if (FName.Contains(CustomFontPrefix.ToLower()))
                        return FontStatus.Replaced;
                    else
                        return FontStatus.Original;
                if (Font == Family)
                    if (FName.Contains(CustomFontPrefix.ToLower()))
                        return FontStatus.Replaced;
                    else
                        return FontStatus.Original;
            }
            return FontStatus.Unknown;
        }

        
    }

    public enum FontStatus {
        Unknown,
        Replaced,
        Redirected,
        Original
    }
    public enum Charset : byte {
        ANSI = 0x00,
        Default = 0x01,
        Symbol = 0x02,
        ShiftJis = 0x80,
        Hangul = 0x81,
        GB2312 = 0x86,
        ChineseBig5 = 0x88,
        Greek = 0xA1,
        Turkish = 0xA2,
        Hebrew = 0xB1,
        Arabic = 0xB2,
        Baltic = 0xBA,
        Russian = 0xCC,
        Thai = 0xDE,
        EE = 0xEE,
        OEM = 0xFF
    }
    internal class FontWorker {
        private bool MonoSpaced;
        private Charset GetCharset;

        private const int LF_FACESIZE = 32;
        private const int LF_FULLFACESIZE = 64;
        private const int FIXED_PITCH = 1;
        private const int TRUETYPE_FONTTYPE = 0x0004;

        private delegate int FONTENUMPROC(ref ENUMLOGFONT lpelf, ref NEWTEXTMETRIC lpntm, uint FontType, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct LOGFONT {
            public int lfHeight;
            public int lfWidth;
            public int lfEscapement;
            public int lfOrientation;
            public int lfWeight;
            public byte lfItalic;
            public byte lfUnderline;
            public byte lfStrikeOut;
            public byte lfCharSet;
            public byte lfOutPrecision;
            public byte lfClipPrecision;
            public byte lfQuality;
            public byte lfPitchAndFamily;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = LF_FACESIZE)]
            public string lfFaceName;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct TEXTMETRIC {
            public int tmHeight;
            public int tmAscent;
            public int tmDescent;
            public int tmInternalLeading;
            public int tmExternalLeading;
            public int tmAveCharWidth;
            public int tmMaxCharWidth;
            public int tmWeight;
            public int tmOverhang;
            public int tmDigitizedAspectX;
            public int tmDigitizedAspectY;
            public char tmFirstChar;
            public char tmLastChar;
            public char tmDefaultChar;
            public char tmBreakChar;
            public byte tmItalic;
            public byte tmUnderlined;
            public byte tmStruckOut;
            public byte tmPitchAndFamily;
            public byte tmCharSet;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct ENUMLOGFONT {
            public LOGFONT elfLogFont;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = LF_FULLFACESIZE)]
            public string elfFullName;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = LF_FACESIZE)]
            public string elfStyle;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct NEWTEXTMETRIC {
            public int tmHeight;
            public int tmAscent;
            public int tmDescent;
            public int tmInternalLeading;
            public int tmExternalLeading;
            public int tmAveCharWidth;
            public int tmMaxCharWidth;
            public int tmWeight;
            public int tmOverhang;
            public int tmDigitizedAspectX;
            public int tmDigitizedAspectY;
            public char tmFirstChar;
            public char tmLastChar;
            public char tmDefaultChar;
            public char tmBreakChar;
            public byte tmItalic;
            public byte tmUnderlined;
            public byte tmStruckOut;
            public byte tmPitchAndFamily;
            public byte tmCharSet;
            public uint ntmFlags;
            public uint ntmSizeEM;
            public uint ntmCellHeight;
            public uint ntmAvgWidth;
        }

        private string[] Result;
        [DllImport("gdi32.dll", CharSet = CharSet.Auto)]
        private static extern int EnumFontFamiliesEx(IntPtr hdc, ref LOGFONT lpLogfont, FONTENUMPROC lpEnumFontFamExProc, IntPtr lParam, uint dwFlags);
        public string[] GetFonts(Charset FontCharset, bool MonospacedOnly) {
            Result = new string[0];
            GetCharset = FontCharset;
            MonoSpaced = MonospacedOnly;
            Graphics graphics = Graphics.FromHwnd(IntPtr.Zero);
            IntPtr hdc = graphics.GetHdc();
            var logfont = new LOGFONT() { lfCharSet = (byte)Charset.Default };
            EnumFontFamiliesEx(hdc, ref logfont, new FONTENUMPROC(EnumFontFamExProc), IntPtr.Zero, 0);
            graphics.ReleaseHdc();
            return Result;
        }

        private int EnumFontFamExProc(ref ENUMLOGFONT lpelf, ref NEWTEXTMETRIC lpntm, uint FontType, IntPtr lParam) {
            if ((MonoSpaced ? (lpelf.elfLogFont.lfPitchAndFamily & 0x3) == FIXED_PITCH : true) && lpelf.elfLogFont.lfCharSet == (byte)GetCharset) {
                Array.Resize(ref Result, Result.Length + 1);
                Result[Result.Length - 1] = lpelf.elfLogFont.lfFaceName;
            }
            return 1;
        }
    }
}
