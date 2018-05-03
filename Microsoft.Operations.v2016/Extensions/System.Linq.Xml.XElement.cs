using System;
using System.Xml;
using System.Xml.Linq;

namespace Microsoft.Operations
{
    static public partial class ExtendXElement
    {
        /// <summary>
        /// from: TEAMFOUNDATIONSERVER XML NOTIFICATIONS Default fetch is from 'StringFields'.
        /// </summary>
        public static string GetCoreFieldValue(this XmlElement eventData, string section, string target, string referenceName)
        {
            return eventData.GetCoreFieldValue(section, target, referenceName, "StringFields");
        }

        /// <summary>
        /// from: TEAMFOUNDATIONSERVER XML NOTIFICATIONS Extracts a text value from some of the core
        /// fields of the work item, assuming it exists. which comes from "CoreFields" section, where
        /// the data is always present in every packet.
        /// NOTE: In theory some of these fields (like WorkItemID) never change, even though
        ///       'OldValue' and 'NewValue' are both present, with the exception of where the field
        /// value (in TFS) doesn't actually HAVE a value (e.g. 'Assigned To' is blank). For fields
        /// where the data exists but has not changed, the 'OldValue' and 'NewValue' text will be the same.
        /// TODO: Change the 'section' and 'target' to enumerators - that would be a bit cleaner.
        /// </summary>
        /// <example>
        /// System.Id, System.Rev, System.AreaId, System.WorkItemType, System.Title, System.AreaPath,
        /// System.State, System.Reason, System.AssignedTo, System.ChangedBy, System.CreatedBy,
        /// System.ChangedDate, System.CreatedDate, System.AuthorizedAs, System.IterationPath,
        /// </example>
        /// <param name="section">"CoreFields" or "ChangedFields" (only)</param>
        /// <param name="target">
        /// String values .... 'OldValue' or 'NewValue' only, otherwise you might get an error.
        /// </param>
        /// <param name="fieldDataType">Specify 'StringFields' or 'IntegerFields'</param>
        public static string GetCoreFieldValue(this XmlElement eventData, string section, string target, string referenceName, string fieldDataType)
        {
            String referenceNameMatch = string.Format("{0}/{3}/Field[ReferenceName='{1}']/{2}", section, referenceName, target, fieldDataType);
            XmlNode nodeToFind;

            // NOTE: An example of the Pattern we want to find is eventData.SelectSingleNode("CoreFields/StringFields/Field[System.Title]/OldValue);

            nodeToFind = eventData.SelectSingleNode(referenceNameMatch);

            if (nodeToFind != null)
            {
                // It was found, manipulate it.
                return nodeToFind.InnerText;
            }
            else
            {
                // It was not found.
                return string.Empty;
            }
        }

        /// <summary>
        /// from: TEAMFOUNDATIONSERVER XML NOTIFICATIONS Examines the (string) contents of a
        /// specified (string) field and returns true if it appears the item has changed. i.e.
        /// between NewValue and OldValue. NOTE: If the item can't be found then it will return false.
        /// </summary>
        /// <param name="section">"CoreFields" or "ChangedFields" (only)</param>
        public static bool IsFieldModified(this XmlElement eventData, string section, string referenceName)
        {
            // TODO: Eventually this might be condensed into a common routine, or there might also be
            // a TeamFoundationServer routine which is a bit more efficient, which could be used instead.

            bool isChanged = false;

            if (!string.IsNullOrEmpty(referenceName))
            {
                String fieldLocationOld = string.Format("{0}/StringFields/Field[ReferenceName='{1}']/OldValue", section, referenceName);
                String fieldLocationNew = string.Format("{0}/StringFields/Field[ReferenceName='{1}']/NewValue", section, referenceName);

                XmlNode nodeToFindOld = eventData.SelectSingleNode(fieldLocationOld);
                XmlNode nodeToFindNew = eventData.SelectSingleNode(fieldLocationNew);

                if (nodeToFindOld != null && nodeToFindNew != null)
                {
                    if (nodeToFindOld.InnerText != nodeToFindNew.InnerText)
                    {
                        return true; // Nodes were found, result comes from direct comparison of the two values.
                    }
                    else
                    {
                        return false; // the values were there, but they appear to be the same.
                    }
                }
                else
                {
                    return false; // The nodes could not be found, therefore no 'changed' comparison.
                }
            }

            return isChanged;
        }

        public static XElement NullSafeElement(this XElement element, XName name)
        {
            return element == null ? null : element.Element(name);
        }

        public static string NullSafeElementValue(this XElement element, XName name)
        {
            return element == null ? null : (element.Element(name) == null ? null : element.Element(name).Value);
        }
    }
}