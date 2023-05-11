using System;
using System.CodeDom.Compiler;
using System.Configuration;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace ReportMailer
{

    [CompilerGenerated]
    [GeneratedCode("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "15.9.0.0")]
    internal sealed class Posta : ApplicationSettingsBase
    {

        public static Posta Default
        {
            get
            {
                return Posta.defaultInstance;
            }
        }

        [UserScopedSetting]
        [DebuggerNonUserCode]
        [DefaultSettingValue("")]
        public string Mail_Adresi
        {
            get
            {
                return (string)this["Mail_Adresi"];
            }
            set
            {
                this["Mail_Adresi"] = value;
            }
        }

        [UserScopedSetting]
        [DebuggerNonUserCode]
        [DefaultSettingValue("")]
        public string Kullanici
        {
            get
            {
                return (string)this["Kullanici"];
            }
            set
            {
                this["Kullanici"] = value;
            }
        }


        [UserScopedSetting]
        [DebuggerNonUserCode]
        [DefaultSettingValue("")]
        public string Sifre
        {
            get
            {
                return (string)this["Sifre"];
            }
            set
            {
                this["Sifre"] = value;
            }
        }

        [UserScopedSetting]
        [DebuggerNonUserCode]
        [DefaultSettingValue("")]
        public string Pop3
        {
            get
            {
                return (string)this["Pop3"];
            }
            set
            {
                this["Pop3"] = value;
            }
        }

        [UserScopedSetting]
        [DebuggerNonUserCode]
        [DefaultSettingValue("")]
        public string Pop3_Port
        {
            get
            {
                return (string)this["Pop3_Port"];
            }
            set
            {
                this["Pop3_Port"] = value;
            }
        }


        [UserScopedSetting]
        [DebuggerNonUserCode]
        [DefaultSettingValue("")]
        public string Smtp
        {
            get
            {
                return (string)this["Smtp"];
            }
            set
            {
                this["Smtp"] = value;
            }
        }

        [UserScopedSetting]
        [DebuggerNonUserCode]
        [DefaultSettingValue("")]
        public string Smtp_Port
        {
            get
            {
                return (string)this["Smtp_Port"];
            }
            set
            {
                this["Smtp_Port"] = value;
            }
        }

        [UserScopedSetting]
        [DebuggerNonUserCode]
        [DefaultSettingValue("")]
        public string Rapor_Saati
        {
            get
            {
                return (string)this["Rapor_Saati"];
            }
            set
            {
                this["Rapor_Saati"] = value;
            }
        }

        [UserScopedSetting]
        [DebuggerNonUserCode]
        [DefaultSettingValue("0")]
        public int Pazartesi
        {
            get
            {
                return (int)this["Pazartesi"];
            }
            set
            {
                this["Pazartesi"] = value;
            }
        }

        [UserScopedSetting]
        [DebuggerNonUserCode]
        [DefaultSettingValue("0")]
        public int Sali
        {
            get
            {
                return (int)this["Sali"];
            }
            set
            {
                this["Sali"] = value;
            }
        }


        [UserScopedSetting]
        [DebuggerNonUserCode]
        [DefaultSettingValue("0")]
        public int Çarşamba
        {
            get
            {
                return (int)this["Çarşamba"];
            }
            set
            {
                this["Çarşamba"] = value;
            }
        }


        [UserScopedSetting]
        [DebuggerNonUserCode]
        [DefaultSettingValue("0")]
        public int Perşembe
        {
            get
            {
                return (int)this["Perşembe"];
            }
            set
            {
                this["Perşembe"] = value;
            }
        }


        [UserScopedSetting]
        [DebuggerNonUserCode]
        [DefaultSettingValue("0")]
        public int Cuma
        {
            get
            {
                return (int)this["Cuma"];
            }
            set
            {
                this["Cuma"] = value;
            }
        }


        [UserScopedSetting]
        [DebuggerNonUserCode]
        [DefaultSettingValue("0")]
        public int Cumartesi
        {
            get
            {
                return (int)this["Cumartesi"];
            }
            set
            {
                this["Cumartesi"] = value;
            }
        }


        [UserScopedSetting]
        [DebuggerNonUserCode]
        [DefaultSettingValue("0")]
        public int Pazar
        {
            get
            {
                return (int)this["Pazar"];
            }
            set
            {
                this["Pazar"] = value;
            }
        }


        [UserScopedSetting]
        [DebuggerNonUserCode]
        [DefaultSettingValue("")]
        public string SQL_Server
        {
            get
            {
                return (string)this["SQL_Server"];
            }
            set
            {
                this["SQL_Server"] = value;
            }
        }


        [UserScopedSetting]
        [DebuggerNonUserCode]
        [DefaultSettingValue("")]
        public string SQL_Database
        {
            get
            {
                return (string)this["SQL_Database"];
            }
            set
            {
                this["SQL_Database"] = value;
            }
        }


        [UserScopedSetting]
        [DebuggerNonUserCode]
        [DefaultSettingValue("")]
        public string SQL_User
        {
            get
            {
                return (string)this["SQL_User"];
            }
            set
            {
                this["SQL_User"] = value;
            }
        }

        [UserScopedSetting]
        [DebuggerNonUserCode]
        [DefaultSettingValue("")]
        public string SQL_Password
        {
            get
            {
                return (string)this["SQL_Password"];
            }
            set
            {
                this["SQL_Password"] = value;
            }
        }




        private static Posta defaultInstance = (Posta)SettingsBase.Synchronized(new Posta());
    }
}
