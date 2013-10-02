using System.Windows;
using System.Reflection;
namespace B1C.Installer
{
    /// <summary>
    /// Instance of the InstallerInfo class.
    /// </summary>
    /// <author>Frank Alunni</author>
    public class InstallerInfo
    {
        /// <summary>
        /// Add on group
        /// </summary>
        private string addOnGroup = "M";

        /// <summary>
        /// Estimated installation time
        /// </summary>
        private string estInstTime = "30";

        /// <summary>
        /// estimated unistall time
        /// </summary>
        private string estUninstTime = "30";

        /// <summary>
        /// Partner contact data
        /// </summary>
        private string partnerContactData = "www.illumiti.com"; // Partner Contact Data

        /// <summary>
        /// Partner Name
        /// </summary>
        private string partnerName = "Illumiti"; // Partner name

        /// <summary>
        /// Partner Namespace
        /// </summary>
        private string partnerNamespace = "XXX"; // Partner Namespace

        /// <summary>
        /// Unistall parameter
        /// </summary>
        private string uninstallParameters = "/U";

        /// <summary>
        /// Initializes a new instance of the <see cref="InstallerInfo"/> class.
        /// </summary>
        public InstallerInfo()
        {
            System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();

            object[] attributes = assembly.GetCustomAttributes(true);

            foreach (object attribute in attributes)
            {
                if (attribute is System.Reflection.AssemblyProductAttribute)
                {
                    this.AddOnExeFile = ((System.Reflection.AssemblyProductAttribute)attribute).Product;
                }
            }

            foreach (object attribute in attributes)
            {
                if (attribute is System.Reflection.AssemblyTitleAttribute)
                {
                    this.ApplicationName = ((System.Reflection.AssemblyTitleAttribute)attribute).Title;
                }
            }

            foreach (object attribute in attributes)
            {
                if (attribute is System.Reflection.AssemblyFileVersionAttribute)
                {
                    this.ApplicationVersion = ((System.Reflection.AssemblyFileVersionAttribute)attribute).Version;
                }
            }

            foreach (object attribute in attributes)
            {
                if (attribute is System.Reflection.AssemblyCompanyAttribute)
                {
                    this.PartnerName = ((System.Reflection.AssemblyCompanyAttribute)attribute).Company;
                }
            }


            //Load Support Files
            Assembly thisExe = Assembly.GetExecutingAssembly();
            string nameNamespace = this.GetType().ToString();
            nameNamespace = this.GetType().Namespace + ".Resources";

            string[] allEmbeddedResources = thisExe.GetManifestResourceNames();
            var addonResources = new System.Collections.ArrayList();

            foreach (string allEmbeddedResource in allEmbeddedResources)
            {
                if (allEmbeddedResource.StartsWith(nameNamespace))
                {
                    addonResources.Add(allEmbeddedResource);   
                }
            }

            this.AddonFileNames = addonResources.ToArray();

            nameNamespace = this.GetType().Namespace + ".SharedResources";
            addonResources = new System.Collections.ArrayList();
            foreach (string allEmbeddedResource in allEmbeddedResources)
            {
                if (allEmbeddedResource.StartsWith(nameNamespace))
                {
                    addonResources.Add(allEmbeddedResource);
                }
            }

            this.AddonSharedFileNames = addonResources.ToArray();
        }

        /// <summary>
        /// Gets or sets AddOnGroup.
        /// </summary>
        /// <value>
        /// The add on group.
        /// </value>
        public string AddOnGroup
        {
            get
            {
                return this.addOnGroup;
            }

            set
            {
                this.addOnGroup = value;
            }
        }

        /// <summary>
        /// Gets or sets ApplicationName.
        /// </summary>
        /// <value>The application name.</value>
        public string ApplicationName
        {
            get;

            set;
        }

        /// <summary>
        /// Gets or sets ApplicationVersion.
        /// </summary>
        /// <value>
        /// The application version.
        /// </value>
        public string ApplicationVersion
        {
            get;

            set;
        }

        /// <summary>
        /// Gets or sets EstInstTime.
        /// </summary>
        /// <value>
        /// The est inst time.
        /// </value>
        public string EstInstTime
        {
            get
            {
                return this.estInstTime;
            }

            set
            {
                this.estInstTime = value;
            }
        }

        /// <summary>
        /// Gets or sets EstUninstTime.
        /// </summary>
        /// <value>
        /// The est uninst time.
        /// </value>
        public string EstUninstTime
        {
            get
            {
                return this.estUninstTime;
            }

            set
            {
                this.estUninstTime = value;
            }
        }

        /// <summary>
        /// Gets or sets InstallParameters.
        /// </summary>
        /// <value>
        /// The install parameters.
        /// </value>
        public string InstallParameters
        {
            get;

            set;
        }

        /// <summary>
        /// Gets or sets PartnerContactData.
        /// </summary>
        /// <value>
        /// The partner contact data.
        /// </value>
        public string PartnerContactData
        {
            get
            {
                return this.partnerContactData;
            }

            set
            {
                this.partnerContactData = value;
            }
        }

        /// <summary>
        /// Gets or sets PartnerName.
        /// </summary>
        /// <value>
        /// The partner name.
        /// </value>
        public string PartnerName
        {
            get
            {
                return this.partnerName;
            }

            set
            {
                this.partnerName = value;
            }
        }

        /// <summary>
        /// Gets or sets PartnerNamespace.
        /// </summary>
        /// <value>
        /// The partner namespace.
        /// </value>
        public string PartnerNamespace
        {
            get
            {
                return this.partnerNamespace;
            }

            set
            {
                this.partnerNamespace = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether RestartNeeded.
        /// </summary>
        /// <value>
        /// The restart needed.
        /// </value>
        public bool RestartNeeded
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets SupportFiles.
        /// </summary>
        /// <value>
        /// The support files.
        /// </value>
        public object[] AddonFileNames
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets SupportFiles.
        /// </summary>
        /// <value>
        /// The support files.
        /// </value>
        public object[] AddonSharedFileNames
        {
            get;
            set;
        }


        /// <summary>
        /// Gets or sets AddOnExeFile.
        /// </summary>
        /// <value>The add on exe file.</value>
        public string AddOnExeFile
        {
            get;

            set;
        }

        /// <summary>
        /// Gets or sets UninstallParameters.
        /// </summary>
        /// <value>
        /// The uninstall parameters.
        /// </value>
        public string UninstallParameters
        {
            get
            {
                return this.uninstallParameters;
            }

            set
            {
                this.uninstallParameters = value;
            }
        }
    }
}
