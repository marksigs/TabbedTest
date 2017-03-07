using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.Web;
using System.Web.Configuration;

namespace MSC.IntegrationService.ConfigService
{
    [Serializable]
    public class ConfigManager
    {
        public ConfigManager()
        {
            //constructor
        }
        public MessageConfig GetMessageConfig(string configSection)
        {
            MessageConfig config;
            try
            {
                if (HttpContext.Current != null)
                {
                    config = (MessageConfig)WebConfigurationManager.GetSection(configSection);
                }
                else
                {
                    config = (MessageConfig)ConfigurationManager.GetSection(configSection);
                }
                return config;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
