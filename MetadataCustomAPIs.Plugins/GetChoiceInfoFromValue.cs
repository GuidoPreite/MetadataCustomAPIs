using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace MetadataCustomAPIs.Plugins
{
    public class GetChoiceInfoFromValue : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            if (context.MessageName.Equals("mca_GetChoiceInfoFromValue") && context.Stage.Equals(30))
            {
                try
                {
                    IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                    IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                    ParameterCollection inputParameters = context.InputParameters;
                    ParameterCollection outputParameters = context.OutputParameters;

                    string tableName = inputParameters["TableName"] as string;
                    string columnName = inputParameters["ColumnName"] as string;
                    int? value = inputParameters["Value"] as int?;
                    int? languageCode = inputParameters["LanguageCode"] as int?;

                    // Default Output Parameters
                    bool choiceFound = false;
                    string choiceLabel = "";
                    string choiceColor = "";
                    string choiceExternalValue = "";
                    string allLanguageLabels = "";
                    int choiceLanguageCode = -1;

                    bool canProceed = true;
                    if (!string.IsNullOrWhiteSpace(tableName)) { tableName = tableName.ToLower(); } else { canProceed = false; }
                    if (!string.IsNullOrWhiteSpace(columnName)) { columnName = columnName.ToLower(); } else { canProceed = false; }

                    if (value == null) { canProceed = false; }

                    if (canProceed == true)
                    {
                        RetrieveAttributeRequest retrieveAttributeRequest = new RetrieveAttributeRequest
                        {
                            EntityLogicalName = tableName,
                            LogicalName = columnName,
                            RetrieveAsIfPublished = true
                        };

                        RetrieveAttributeResponse retrieveAttributeResponse = (RetrieveAttributeResponse)service.Execute(retrieveAttributeRequest);
                        EnumAttributeMetadata attributeMetadata = (EnumAttributeMetadata)retrieveAttributeResponse.AttributeMetadata;

                        if (attributeMetadata.AttributeType == AttributeTypeCode.Picklist || attributeMetadata.AttributeType == AttributeTypeCode.Virtual)
                        {
                            if (attributeMetadata.OptionSet != null)
                            {
                                OptionMetadataCollection options = attributeMetadata.OptionSet.Options;
                                foreach (OptionMetadata option in options)
                                {
                                    if (option.Value != null && option.Value.Value == value.Value)
                                    {
                                        choiceFound = true;
                                        choiceColor = option.Color;
                                        choiceExternalValue = option.ExternalValue;

                                        if (option.Label != null && option.Label.UserLocalizedLabel != null)
                                        {
                                            choiceLabel = option.Label.UserLocalizedLabel.Label;
                                            choiceLanguageCode = option.Label.UserLocalizedLabel.LanguageCode;
                                        }

                                        if (option.Label != null && option.Label.LocalizedLabels != null)
                                        {
                                            foreach (LocalizedLabel localizedLabel in option.Label.LocalizedLabels)
                                            {
                                                string currentLabel = localizedLabel.Label;
                                                int currentLanguageCode = localizedLabel.LanguageCode;

                                                allLanguageLabels += $"\"LanguageCode\": {currentLanguageCode}, \"Label\": \"{currentLabel}\",";

                                                if (languageCode.HasValue && languageCode != 0 && localizedLabel.LanguageCode == languageCode.Value)
                                                {
                                                    choiceLabel = currentLabel;
                                                    choiceLanguageCode = currentLanguageCode;
                                                }
                                            }
                                        }
                                    }
                                    if (choiceFound == true) { break; }
                                }
                            }
                        }
                    }

                    if (string.IsNullOrWhiteSpace(allLanguageLabels))
                    {
                        allLanguageLabels = "[]";
                    }
                    else
                    {
                        allLanguageLabels = "[" + allLanguageLabels.Substring(0, allLanguageLabels.Length - 1) + "]";
                    }

                    context.OutputParameters["ChoiceFound"] = choiceFound;
                    context.OutputParameters["ChoiceLabel"] = choiceLabel;
                    context.OutputParameters["ChoiceColor"] = choiceColor;
                    context.OutputParameters["ChoiceExternalValue"] = choiceExternalValue;
                    context.OutputParameters["ChoiceAllLanguageLabels"] = allLanguageLabels;
                    context.OutputParameters["ChoiceLanguageCode"] = choiceLanguageCode;
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException($"GetChoiceInfoFromValue Error. Details: {ex.Message}");
                }
            }
        }
    }
}