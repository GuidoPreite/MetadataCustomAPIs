using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace MetadataCustomAPIs.Plugins
{
    public class GetChoiceInfoFromLabel : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            if (context.MessageName.Equals("mca_GetChoiceInfoFromLabel") && context.Stage.Equals(30))
            {
                try
                {
                    IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                    IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                    ParameterCollection inputParameters = context.InputParameters;
                    ParameterCollection outputParameters = context.OutputParameters;

                    string tableName = inputParameters["TableName"] as string;
                    string columnName = inputParameters["ColumnName"] as string;
                    string label = inputParameters["Label"] as string;
                    bool? labelCaseSensitive = inputParameters["LabelCaseSensitive"] as bool?;
                    int? languageCode = inputParameters["LanguageCode"] as int?;

                    StringComparison comparison = StringComparison.OrdinalIgnoreCase;
                    if (labelCaseSensitive != null && labelCaseSensitive.Value == true) { comparison = StringComparison.Ordinal; }

                    // Default Output Parameters
                    bool choiceFound = false;
                    int choiceValue = -1;
                    string choiceColor = "";
                    string choiceExternalValue = "";
                    int choiceLanguageCode = -1;

                    bool canProceed = true;
                    if (!string.IsNullOrWhiteSpace(tableName)) { tableName = tableName.ToLower(); } else { canProceed = false; }
                    if (!string.IsNullOrWhiteSpace(columnName)) { columnName = columnName.ToLower(); } else { canProceed = false; }
                    if (!string.IsNullOrWhiteSpace(label)) { label = label.ToLower(); } else { canProceed = false; }

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
                                    string currentLabel = "";
                                    if (option.Label != null && option.Label.LocalizedLabels != null)
                                    {
                                        foreach (LocalizedLabel localizedLabel in option.Label.LocalizedLabels)
                                        {
                                            if (languageCode.HasValue && languageCode != 0 && localizedLabel.LanguageCode != languageCode.Value) { continue; }
                                            currentLabel = localizedLabel.Label;

                                            if (currentLabel.Equals(label, comparison))
                                            {
                                                choiceFound = true;
                                                choiceValue = option.Value.GetValueOrDefault(-1);
                                                choiceColor = option.Color;
                                                choiceExternalValue = option.ExternalValue;
                                                choiceLanguageCode = localizedLabel.LanguageCode;
                                                break;
                                            }
                                        }
                                        if (choiceFound == true) { break; }
                                    }
                                }
                            }
                        }
                    }
                    context.OutputParameters["ChoiceFound"] = choiceFound;
                    context.OutputParameters["ChoiceValue"] = choiceValue;
                    context.OutputParameters["ChoiceColor"] = choiceColor;
                    context.OutputParameters["ChoiceExternalValue"] = choiceExternalValue;
                    context.OutputParameters["ChoiceLanguageCode"] = choiceLanguageCode;
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException($"GetChoiceInfoFromLabel Error. Details: {ex.Message}");
                }
            }
        }
    }
}