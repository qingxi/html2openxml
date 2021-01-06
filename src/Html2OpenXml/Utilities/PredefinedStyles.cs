﻿using System.Resources;

namespace HtmlToOpenXml
{
    /// <summary>
    /// Helper class to get chunks of OpenXml predefined style.
    /// </summary>
    internal class PredefinedStyles
    {
        private static global::System.Resources.ResourceManager resourceMan;


        /// <summary>
        /// Retrieves the embedded resource.
        /// </summary>
        /// <param name="styleName">The key name of the resource to find.</param>
        public static string GetOuterXml(string styleName)
        {
            return ResourceManager.GetString(styleName);
        }


        /// <summary>
        /// Returns the cached ResourceManager instance used by this class.
        /// </summary>
        private static ResourceManager ResourceManager
        {
            get
            {
                if (object.ReferenceEquals(resourceMan, null))
                {
                    ResourceManager temp = new ResourceManager("HtmlToOpenXml.PredefinedStyles",
                        typeof(PredefinedStyles).Assembly);

                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
    }
}
