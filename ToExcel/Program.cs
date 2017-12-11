using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using YamlDotNet.Serialization;

namespace ToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            string outputPath = Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory.TrimEnd('\\'))), "Agencies.xlsx");
            const string allFilesPath = @"C:\Code\Open-Water-Rate-Specification\full_utility_rates\California";

            List<string> fileNames = Directory.GetDirectories(allFilesPath).Select(x => Directory.GetFiles(x).FirstOrDefault()).Where(x => x != null).ToList();

            List<AgencyObject> objects = new List<AgencyObject>();

            foreach (var fileName in fileNames)
            {
                AgencyObject agencyObject = new AgencyObject();
                objects.Add(agencyObject);
                using (StringReader reader = new StringReader(File.ReadAllText(fileName)))
                {
                    var deserializer = new DeserializerBuilder().Build();
                    var yamlObject = deserializer.Deserialize(reader);

                    var serializer = new SerializerBuilder().JsonCompatible().Build();
                    var json = serializer.Serialize(yamlObject);

                    JObject jItems = JObject.Parse(json);
                    foreach (KeyValuePair<string, JToken> jItem in jItems)
                    {
                        if (jItem.Key.Equals("metadata", StringComparison.InvariantCultureIgnoreCase))
                        {
                            agencyObject.Metadata = ParseMetadata(jItem.Value);
                        }
                        else if (jItem.Key.Equals("rate_structure", StringComparison.InvariantCultureIgnoreCase))
                        {
                            agencyObject.RateStructures = ParseRateStructure(jItem.Value);
                        }
                        else if (jItem.Key.Equals("capacity_charge", StringComparison.InvariantCultureIgnoreCase))
                        {
                            agencyObject.CapacityCharge = ParseAgencyList(jItem.Value);
                        }
                        else
                        {
                        }
                    }
                }
            }
            ExportToExcelFile(outputPath, objects);
        }

        private static AgencyMetadata ParseMetadata(JToken jItem)
        {
            AgencyMetadata metadata = new AgencyMetadata();
            foreach (KeyValuePair<string, JToken> jMetadataItem in (JObject) jItem)
            {
                if (jMetadataItem.Key.Equals("utility_name", StringComparison.InvariantCultureIgnoreCase))
                {
                    metadata.UtilityName = jMetadataItem.Value.ToString();
                }
                else if (jMetadataItem.Key.Equals("effective_date", StringComparison.InvariantCultureIgnoreCase))
                {
                    metadata.EffectiveDate = jMetadataItem.Value.ToString();
                }
                else if (jMetadataItem.Key.Equals("bill_frequency", StringComparison.InvariantCultureIgnoreCase))
                {
                    metadata.BillFrequency = jMetadataItem.Value.ToString();
                }
                else if (jMetadataItem.Key.Equals("bill_unit", StringComparison.InvariantCultureIgnoreCase))
                {
                    metadata.BillUnit = jMetadataItem.Value.ToString();
                }
                else
                {
                }
            }
            return metadata;
        }

        private static List<AgencyRateStructure> ParseRateStructure(JToken jItem)
        {
            List<AgencyRateStructure> rateStructures = new List<AgencyRateStructure>();
            foreach (KeyValuePair<string, JToken> jRateStructureParentItem in (JObject) jItem)
            {
                AgencyRateStructure rateStructure = new AgencyRateStructure { ConsumerClass = jRateStructureParentItem.Key };
                rateStructures.Add(rateStructure);
                bool isJObject = jRateStructureParentItem.Value is JObject;
                if (!isJObject)
                    continue;
                foreach (KeyValuePair<string, JToken> jRateStructureItem in (JObject) jRateStructureParentItem.Value)
                {
                    if (jRateStructureItem.Key.Equals("adjusted_budget", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.AdjustedBudget = jRateStructureItem.Value.ToString();
                    }
                    else if (jRateStructureItem.Key.Equals("bill", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.Bill = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("budget", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.Budget = jRateStructureItem.Value.ToString();
                    }
                    else if (jRateStructureItem.Key.Equals("commodity_price", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.CommodityPrice = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("days_in_period", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.DaysInPeriod = jRateStructureItem.Value.ToString();
                    }
                    else if (jRateStructureItem.Key.Equals("elevation_rate", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.ElevationRate = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("et_amount", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.EtAmount = jRateStructureItem.Value.ToString();
                    }
                    else if (jRateStructureItem.Key.Equals("fixed_water_service", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.FixedWaterService = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("flat_rate", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.FlatRate = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("flat_rate_commodity", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.FlatRateCommodity = jRateStructureItem.Value.ToString();
                    }
                    else if (jRateStructureItem.Key.Equals("flat_rate_drought", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.FlatRateDrought = jRateStructureItem.Value.ToString();
                    }
                    else if (jRateStructureItem.Key.Equals("gpcd", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.Gpcd = jRateStructureItem.Value.ToString();
                    }
                    else if (jRateStructureItem.Key.Equals("hhsize", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.HhSize = jRateStructureItem.Value.ToString();
                    }
                    else if (jRateStructureItem.Key.Equals("indoor", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.Indoor = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("indoor_price", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.IndoorPrice = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("irrigation_efficiency", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.IrrigationEfficiency = jRateStructureItem.Value.ToString();
                    }
                    else if (jRateStructureItem.Key.Equals("mwd_ready_to_serve", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.MwdReadyToServe = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("outdoor", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.Outdoor = jRateStructureItem.Value.ToString();
                    }
                    else if (jRateStructureItem.Key.Equals("outside_city_service_price", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.OutsideCityServicePrice = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("rate", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.Rate = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("recycled_rate", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.RecycledRate = jRateStructureItem.Value.ToString();
                    }
                    else if (jRateStructureItem.Key.Equals("sdcwa_infrastructure", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.SdcwaInfrastructure = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("sewer_price", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.SewerPrice = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("tier_prices", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.TierPrices = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("tier_prices_commodity", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.TierPricesCommodity = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("tier_prices_drought", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.TierPricesDrought = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("tier_rates", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.TierRates = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("tier_starts", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.TierStarts = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("tier_starts_commodity", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.TierStartsCommodity = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("tier_starts_drought", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.TierStartsDrought = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("wrap_discount", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.WrapDiscount = jRateStructureItem.Value.ToString();
                    }
                    else if (jRateStructureItem.Key.Equals("capital_facilities_charge", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.CapitalFacilitiesCharge = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("capital_improvements_charge", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.CapitalImprovementsCharge = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("carw_charge", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.CarwCharge = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("commodity_charge", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.CommodityCharge = jRateStructureItem.Value.ToString();
                    }
                    else if (jRateStructureItem.Key.Equals("connection_fee", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.ConnectionFee = jRateStructureItem.Value.ToString();
                    }
                    else if (jRateStructureItem.Key.Equals("cost_adjustment_charge", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.CostAdjustmentCharge = jRateStructureItem.Value.ToString();
                    }
                    else if (jRateStructureItem.Key.Equals("elevation_charge", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.ElevationCharge = jRateStructureItem.Value.ToString();
                    }
                    else if (jRateStructureItem.Key.Equals("fixed_sewer_charge", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.FixedSewerCharge = jRateStructureItem.Value.ToString();
                    }
                    else if (jRateStructureItem.Key.Equals("fixed_wastewater_charge", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.FixedWastewaterCharge = jRateStructureItem.Value.ToString();
                    }
                    else if (jRateStructureItem.Key.Equals("flat_service_charge", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.FlatServiceCharge = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("minimum_charge", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.MinimumCharge = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("outside_city_service_charge", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.OutsideCityServiceCharge = jRateStructureItem.Value.ToString();
                    }
                    else if (jRateStructureItem.Key.Equals("recycled_charge", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.RecycledCharge = jRateStructureItem.Value.ToString();
                    }
                    else if (jRateStructureItem.Key.Equals("reliability_charge", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.ReliabilityCharge = jRateStructureItem.Value.ToString();
                    }
                    else if (jRateStructureItem.Key.Equals("sanitation_charge", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.SanitationCharge = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("service_charge", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.ServiceCharge = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("sewer_charge", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.SewerCharge = jRateStructureItem.Value.ToString();
                    }
                    else if (jRateStructureItem.Key.Equals("variable_wastewater_charge", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.VariableWastewaterCharge = jRateStructureItem.Value.ToString();
                    }
                    else if (jRateStructureItem.Key.Equals("drought_surcharge", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.DroughtSurcharge = jRateStructureItem.Value.ToString();
                    }
                    else if (jRateStructureItem.Key.Equals("fixed_drought_surcharge", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.FixedDroughtSurcharge = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("safe_drinking_water_surcharge", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.SafeDrinkingWaterSurcharge = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("utility_surcharge", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.UtilitySurcharge = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("utility_surcharge_price", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.UtilitySurchargePrice = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("variable_drought_surcharge", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.VariableDroughtSurcharge = jRateStructureItem.Value.ToString();
                    }
                    else if (jRateStructureItem.Key.Equals("wrap_surcharge", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.WrapSurcharge = jRateStructureItem.Value.ToString();
                    }
                    else if (jRateStructureItem.Key.Equals("landscape_factor", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.LandscapeFactor = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("monthly_plant_factor", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.MonthlyPlantFactor = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else if (jRateStructureItem.Key.Equals("volumetric_conversion_factor", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.VolumetricConversionFactor = jRateStructureItem.Value.ToString();
                    }
                    else if (jRateStructureItem.Key.Equals("weather_adjustment_factor", StringComparison.InvariantCultureIgnoreCase))
                    {
                        rateStructure.WeatherAdjustmentFactor = ParseAgencyList(jRateStructureItem.Value);
                    }
                    else
                    {
                    }
                }
            }
            return rateStructures;
        }

        private static AgencyList ParseAgencyList(JToken jToken)
        {
            AgencyList list = new AgencyList
            {
                Values = new Dictionary<string, List<string>>()
            };

            if (jToken is JValue)
            {
                list.Values["All"] = new List<string> { jToken.ToString() };
            }
            else if (jToken is JArray)
            {
                list.Values["All"] = new List<string>();
                foreach (JValue jArrayItem in (JArray) jToken)
                    list.Values["All"].Add(jArrayItem.ToString());
            }
            else if (jToken is JObject)
            {
                foreach (KeyValuePair<string, JToken> kvp in (JObject) jToken)
                {
                    if (kvp.Key.Equals("depends_on", StringComparison.InvariantCultureIgnoreCase))
                    {
                        list.DependsOn = new List<string>();
                        if (kvp.Value is JValue)
                        {
                            list.DependsOn.Add(kvp.Value.ToString());
                        }
                        else if (kvp.Value is JArray)
                        {
                            foreach (JValue jArrayItem in (JArray) kvp.Value)
                            {
                                list.DependsOn.Add(jArrayItem.ToString());
                            }
                        }
                        else
                        {
                        }
                    }
                    else if (kvp.Key.Equals("values", StringComparison.InvariantCultureIgnoreCase))
                    {
                        if (kvp.Value is JObject)
                        {
                            foreach (KeyValuePair<string, JToken> kvp2 in (JObject) kvp.Value)
                            {
                                list.Values[kvp2.Key] = new List<string>();
                                if (kvp2.Value is JValue)
                                {
                                    list.Values[kvp2.Key].Add(kvp2.Value.ToString());
                                }
                                else if (kvp2.Value is JArray)
                                {
                                    foreach (JValue jValue in (JArray) (kvp2.Value))
                                    {
                                        list.Values[kvp2.Key].Add(jValue.ToString());
                                    }
                                }
                                else
                                {
                                }
                            }
                        }
                        else if (kvp.Value is JArray)
                        {
                            foreach (JObject jArrayItem in (JArray) kvp.Value)
                            {
                                foreach (KeyValuePair<string, JToken> kvp2 in jArrayItem)
                                {
                                    if (kvp2.Value is JValue)
                                    {
                                        list.Values[kvp2.Key] = new List<string> { kvp2.Value.ToString() };
                                    }
                                    else
                                    {
                                    }
                                }
                            }
                        }
                        else
                        {
                        }
                    }
                    else
                    {
                    }
                }
            }
            else
            {
            }

            return list;
        }

        private static void ExportToExcelFile(string outputPath, List<AgencyObject> agencyObjects)
        {
            using (ExcelPackage excel = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add("data");
                ExportToExcelWorksheet(worksheet, agencyObjects);

                FileInfo file = new FileInfo(outputPath);
                excel.SaveAs(file);
            }
        }

        private static void ExportToExcelWorksheet(ExcelWorksheet worksheet, List<AgencyObject> agencyObjects)
        {
            var totalRows = agencyObjects.Count;

            Dictionary<Tuple<string, string, string, string>, Func<AgencyObject, string>> allData = new Dictionary<Tuple<string, string, string, string>, Func<AgencyObject, string>>
    {
        // metadata
        //{Tuple.Create("", "utility name", "", ""), obj => obj.Metadata.UtilityName},
        {Tuple.Create("", "effective_date", "", ""), obj => obj.Metadata.EffectiveDate},
        {Tuple.Create("", "bill_frequency", "", ""), obj => obj.Metadata.BillFrequency},
        {Tuple.Create("", "bill_unit", "", ""), obj => obj.Metadata.BillUnit},
    };

            // capacity charge
            foreach (string capacityKey in GetListKeys(agencyObjects.Select(x => x.CapacityCharge)))
            {
                allData.Add(Tuple.Create("", "capacity_charge", "meter_size", capacityKey),
                    obj => obj.CapacityCharge != null && obj.CapacityCharge.Values.ContainsKey(capacityKey)
                        ? obj.CapacityCharge.Values[capacityKey].FirstOrDefault()
                        : null);
            }

            var allRateStructures = agencyObjects.SelectMany(x => x.RateStructures).Where(x => x != null);

            // rate structure
            foreach (string consumerClass in agencyObjects.SelectMany(x => x.RateStructures.Select(y => y.ConsumerClass)).Distinct().OrderBy(x => x))
            {
                var consumerClassRateStructures = allRateStructures.Where(rs => rs.ConsumerClass == consumerClass).ToList();

                // strings
                allData.Add(Tuple.Create(consumerClass, "adjusted_budget", "", ""), obj => GetRateStructure(obj, consumerClass) != null ? GetRateStructure(obj, consumerClass).AdjustedBudget : null);
                allData.Add(Tuple.Create(consumerClass, "budget", "", ""), obj => GetRateStructure(obj, consumerClass) != null ? GetRateStructure(obj, consumerClass).Budget : null);
                allData.Add(Tuple.Create(consumerClass, "commodity_charge", "", ""), obj => GetRateStructure(obj, consumerClass) != null ? GetRateStructure(obj, consumerClass).CommodityCharge : null);
                allData.Add(Tuple.Create(consumerClass, "connection_fee", "", ""), obj => GetRateStructure(obj, consumerClass) != null ? GetRateStructure(obj, consumerClass).ConnectionFee : null);
                allData.Add(Tuple.Create(consumerClass, "cost_adjustment_charge", "", ""), obj => GetRateStructure(obj, consumerClass) != null ? GetRateStructure(obj, consumerClass).CostAdjustmentCharge : null);
                allData.Add(Tuple.Create(consumerClass, "days_in_period", "", ""), obj => GetRateStructure(obj, consumerClass) != null ? GetRateStructure(obj, consumerClass).DaysInPeriod : null);
                allData.Add(Tuple.Create(consumerClass, "drought_surcharge", "", ""), obj => GetRateStructure(obj, consumerClass) != null ? GetRateStructure(obj, consumerClass).DroughtSurcharge : null);
                allData.Add(Tuple.Create(consumerClass, "elevation_charge", "", ""), obj => GetRateStructure(obj, consumerClass) != null ? GetRateStructure(obj, consumerClass).ElevationCharge : null);
                allData.Add(Tuple.Create(consumerClass, "et_amount", "", ""), obj => GetRateStructure(obj, consumerClass) != null ? GetRateStructure(obj, consumerClass).EtAmount : null);
                allData.Add(Tuple.Create(consumerClass, "fixed_sewer_charge", "", ""), obj => GetRateStructure(obj, consumerClass) != null ? GetRateStructure(obj, consumerClass).FixedSewerCharge : null);
                allData.Add(Tuple.Create(consumerClass, "fixed_wastewater_charge", "", ""), obj => GetRateStructure(obj, consumerClass) != null ? GetRateStructure(obj, consumerClass).FixedWastewaterCharge : null);
                allData.Add(Tuple.Create(consumerClass, "flat_rate_commodity", "", ""), obj => GetRateStructure(obj, consumerClass) != null ? GetRateStructure(obj, consumerClass).FlatRateCommodity : null);
                allData.Add(Tuple.Create(consumerClass, "flat_rate_drought", "", ""), obj => GetRateStructure(obj, consumerClass) != null ? GetRateStructure(obj, consumerClass).FlatRateDrought : null);
                allData.Add(Tuple.Create(consumerClass, "gpcd", "", ""), obj => GetRateStructure(obj, consumerClass) != null ? GetRateStructure(obj, consumerClass).Gpcd : null);
                allData.Add(Tuple.Create(consumerClass, "hhsize", "", ""), obj => GetRateStructure(obj, consumerClass) != null ? GetRateStructure(obj, consumerClass).HhSize : null);
                allData.Add(Tuple.Create(consumerClass, "irrigation_efficiency", "", ""), obj => GetRateStructure(obj, consumerClass) != null ? GetRateStructure(obj, consumerClass).IrrigationEfficiency : null);
                allData.Add(Tuple.Create(consumerClass, "outdoor", "", ""), obj => GetRateStructure(obj, consumerClass) != null ? GetRateStructure(obj, consumerClass).Outdoor : null);
                allData.Add(Tuple.Create(consumerClass, "outside_city_service_charge", "", ""), obj => GetRateStructure(obj, consumerClass) != null ? GetRateStructure(obj, consumerClass).OutsideCityServiceCharge : null);
                allData.Add(Tuple.Create(consumerClass, "recycled_charge", "", ""), obj => GetRateStructure(obj, consumerClass) != null ? GetRateStructure(obj, consumerClass).RecycledCharge : null);
                allData.Add(Tuple.Create(consumerClass, "recycled_rate", "", ""), obj => GetRateStructure(obj, consumerClass) != null ? GetRateStructure(obj, consumerClass).RecycledRate : null);
                allData.Add(Tuple.Create(consumerClass, "reliability_charge", "", ""), obj => GetRateStructure(obj, consumerClass) != null ? GetRateStructure(obj, consumerClass).ReliabilityCharge : null);
                allData.Add(Tuple.Create(consumerClass, "sewer_charge", "", ""), obj => GetRateStructure(obj, consumerClass) != null ? GetRateStructure(obj, consumerClass).SewerCharge : null);
                allData.Add(Tuple.Create(consumerClass, "variable_drought_surcharge", "", ""), obj => GetRateStructure(obj, consumerClass) != null ? GetRateStructure(obj, consumerClass).VariableDroughtSurcharge : null);
                allData.Add(Tuple.Create(consumerClass, "variable_wastewater_charge", "", ""), obj => GetRateStructure(obj, consumerClass) != null ? GetRateStructure(obj, consumerClass).VariableWastewaterCharge : null);
                allData.Add(Tuple.Create(consumerClass, "volumetric_conversion_factor", "", ""), obj => GetRateStructure(obj, consumerClass) != null ? GetRateStructure(obj, consumerClass).VolumetricConversionFactor : null);
                allData.Add(Tuple.Create(consumerClass, "wrap_discount", "", ""), obj => GetRateStructure(obj, consumerClass) != null ? GetRateStructure(obj, consumerClass).WrapDiscount : null);
                allData.Add(Tuple.Create(consumerClass, "wrap_surcharge", "", ""), obj => GetRateStructure(obj, consumerClass) != null ? GetRateStructure(obj, consumerClass).WrapSurcharge : null);

                // lists
                if (consumerClass == "RESIDENTIAL_SINGLE_MOUNTAIN") // depends on wrap_customer
                {
                    allData.Add(Tuple.Create(consumerClass, "bill", "wrap_customer", "Yes"), obj => GetFirstListValue(consumerClass, obj, rs => rs.Bill, "Yes"));
                    allData.Add(Tuple.Create(consumerClass, "bill", "wrap_customer", "No"), obj => GetFirstListValue(consumerClass, obj, rs => rs.Bill, "No"));
                }
                else
                {
                    allData.Add(Tuple.Create(consumerClass, "bill", "", ""), obj => GetFirstListValue(consumerClass, obj, rs => rs.Bill, "All"));
                }
                if (consumerClassRateStructures.Any(rs => rs.CapitalFacilitiesCharge != null))
                {
                    foreach (string key in GetDependentKeys(allRateStructures, rs => rs.CapitalFacilitiesCharge))
                        allData.Add(Tuple.Create(consumerClass, "capital_facilities_charge", "meter_size", key), obj => GetFirstListValue(consumerClass, obj, rs => rs.CapitalFacilitiesCharge, key));
                }
                if (consumerClassRateStructures.Any(rs => rs.CapitalImprovementsCharge != null))
                {
                    foreach (string key in GetDependentKeys(allRateStructures, rs => rs.CapitalImprovementsCharge))
                        allData.Add(Tuple.Create(consumerClass, "capital_improvements_charge", "meter_size", key), obj => GetFirstListValue(consumerClass, obj, rs => rs.CapitalImprovementsCharge, key));
                }
                if (consumerClassRateStructures.Any(rs => rs.CarwCharge != null))
                {
                    foreach (string key in GetDependentKeys(allRateStructures, rs => rs.CarwCharge))
                        allData.Add(Tuple.Create(consumerClass, "carw_charge", "carw_customer|meter_size", key), obj => GetFirstListValue(consumerClass, obj, rs => rs.CarwCharge, key));
                }
                if (consumerClassRateStructures.Any(rs => rs.CommodityPrice != null))
                {
                    foreach (string key in GetDependentKeys(allRateStructures, rs => rs.CommodityPrice))
                        allData.Add(Tuple.Create(consumerClass, "commodity_price", "water_type", key), obj => GetFirstListValue(consumerClass, obj, rs => rs.CommodityPrice, key));
                }
                if (consumerClassRateStructures.Any(rs => rs.FixedWaterService != null))
                {
                    foreach (string key in GetDependentKeys(allRateStructures, rs => rs.FixedWaterService))
                        allData.Add(Tuple.Create(consumerClass, "fixed_water_service", "meter_size", key), obj => GetFirstListValue(consumerClass, obj, rs => rs.FixedWaterService, key));
                }
                if (consumerClassRateStructures.Any(rs => rs.FlatServiceCharge != null))
                {
                    foreach (string key in GetDependentKeys(allRateStructures, rs => rs.FlatServiceCharge))
                        allData.Add(Tuple.Create(consumerClass, "flat_rate_service_charge", "meter_size", key), obj => GetFirstListValue(consumerClass, obj, rs => rs.FlatServiceCharge, key));
                }
                if (consumerClassRateStructures.Any(rs => rs.IndoorPrice != null))
                {
                    foreach (string key in GetDependentKeys(allRateStructures, rs => rs.IndoorPrice))
                        allData.Add(Tuple.Create(consumerClass, "indoor_price", "season", key), obj => GetFirstListValue(consumerClass, obj, rs => rs.IndoorPrice, key));
                }
                if (consumerClassRateStructures.Any(rs => rs.MinimumCharge != null))
                {
                    foreach (string key in GetDependentKeys(allRateStructures, rs => rs.MinimumCharge))
                        allData.Add(Tuple.Create(consumerClass, "minimum_charge", "meter_size", key), obj => GetFirstListValue(consumerClass, obj, rs => rs.MinimumCharge, key));
                }
                if (consumerClassRateStructures.Any(rs => rs.MonthlyPlantFactor != null))
                {
                    foreach (string key in GetDependentKeys(allRateStructures, rs => rs.MonthlyPlantFactor))
                        allData.Add(Tuple.Create(consumerClass, "monthly_plant_factor", "usage_month", key), obj => GetFirstListValue(consumerClass, obj, rs => rs.MonthlyPlantFactor, key));
                }
                if (consumerClassRateStructures.Any(rs => rs.MwdReadyToServe != null))
                {
                    foreach (string key in GetDependentKeys(allRateStructures, rs => rs.MwdReadyToServe))
                        allData.Add(Tuple.Create(consumerClass, "mwd_ready_to_serve", "meter_size", key), obj => GetFirstListValue(consumerClass, obj, rs => rs.MwdReadyToServe, key));
                }
                if (consumerClassRateStructures.Any(rs => rs.OutsideCityServicePrice != null))
                {
                    foreach (string key in GetDependentKeys(allRateStructures, rs => rs.OutsideCityServicePrice))
                        allData.Add(Tuple.Create(consumerClass, "outside_city_service_price", "city_limits", key), obj => GetFirstListValue(consumerClass, obj, rs => rs.OutsideCityServicePrice, key));
                }
                if (consumerClassRateStructures.Any(rs => rs.Rate != null))
                {
                    foreach (string key in GetDependentKeys(allRateStructures, rs => rs.Rate))
                        allData.Add(Tuple.Create(consumerClass, "rate", "water_type", key), obj => GetFirstListValue(consumerClass, obj, rs => rs.Rate, key));
                }
                if (consumerClassRateStructures.Any(rs => rs.SafeDrinkingWaterSurcharge != null))
                {
                    foreach (string key in GetDependentKeys(allRateStructures, rs => rs.SafeDrinkingWaterSurcharge))
                        allData.Add(Tuple.Create(consumerClass, "safe_drinking_water_surcharge", "meter_size", key), obj => GetFirstListValue(consumerClass, obj, rs => rs.SafeDrinkingWaterSurcharge, key));
                }
                if (consumerClassRateStructures.Any(rs => rs.SdcwaInfrastructure != null))
                {
                    foreach (string key in GetDependentKeys(allRateStructures, rs => rs.SdcwaInfrastructure))
                        allData.Add(Tuple.Create(consumerClass, "sdcwa_infrastructure", "meter_size", key), obj => GetFirstListValue(consumerClass, obj, rs => rs.SdcwaInfrastructure, key));
                }
                if (consumerClassRateStructures.Any(rs => rs.SewerPrice != null))
                {
                    foreach (string key in GetDependentKeys(allRateStructures, rs => rs.SewerPrice))
                        allData.Add(Tuple.Create(consumerClass, "sewer_price", "rate_class", key), obj => GetFirstListValue(consumerClass, obj, rs => rs.SewerPrice, key));
                }
                if (consumerClassRateStructures.Any(rs => rs.UtilitySurchargePrice != null))
                {
                    foreach (string key in GetDependentKeys(allRateStructures, rs => rs.UtilitySurchargePrice))
                        allData.Add(Tuple.Create(consumerClass, "utility_surcharge_price", "pressure_zone", key), obj => GetFirstListValue(consumerClass, obj, rs => rs.UtilitySurchargePrice, key));
                }
                if (consumerClassRateStructures.Any(rs => rs.WeatherAdjustmentFactor != null))
                {
                    foreach (string key in GetDependentKeys(allRateStructures, rs => rs.WeatherAdjustmentFactor))
                        allData.Add(Tuple.Create(consumerClass, "weather_adjustment_factor", "usage_zone", key), obj => GetFirstListValue(consumerClass, obj, rs => rs.WeatherAdjustmentFactor, key));
                }

                // difficult lists
                Func<AgencyRateStructure, AgencyList> fnGetList = rs => rs.ElevationRate;
                AddAllMultiDependentDataKeys(allData, consumerClassRateStructures, fnGetList, consumerClass, "elevation_rate");

                fnGetList = rs => rs.FixedDroughtSurcharge;
                AddAllMultiDependentDataKeys(allData, consumerClassRateStructures, fnGetList, consumerClass, "fixed_drought_surcharge");

                fnGetList = rs => rs.FlatRate;
                AddAllMultiDependentDataKeys(allData, consumerClassRateStructures, fnGetList, consumerClass, "flat_rate");

                fnGetList = rs => rs.Indoor;
                AddAllMultiDependentDataKeys(allData, consumerClassRateStructures, fnGetList, consumerClass, "indoor");

                fnGetList = rs => rs.LandscapeFactor;
                AddAllMultiDependentDataKeys(allData, consumerClassRateStructures, fnGetList, consumerClass, "landscape_factor");

                fnGetList = rs => rs.SanitationCharge;
                AddAllMultiDependentDataKeys(allData, consumerClassRateStructures, fnGetList, consumerClass, "sanitation_charge");

                fnGetList = rs => rs.ServiceCharge;
                AddAllMultiDependentDataKeys(allData, consumerClassRateStructures, fnGetList, consumerClass, "service_charge");

                fnGetList = rs => rs.TierPrices;
                AddAllMultiDependentDataKeys(allData, consumerClassRateStructures, fnGetList, consumerClass, "tier_prices");

                fnGetList = rs => rs.TierPricesCommodity;
                AddAllMultiDependentDataKeys(allData, consumerClassRateStructures, fnGetList, consumerClass, "tier_prices_commodity");

                fnGetList = rs => rs.TierPricesDrought;
                AddAllMultiDependentDataKeys(allData, consumerClassRateStructures, fnGetList, consumerClass, "tier_prices_drought");

                fnGetList = rs => rs.TierRates;
                AddAllMultiDependentDataKeys(allData, consumerClassRateStructures, fnGetList, consumerClass, "tier_rates");

                fnGetList = rs => rs.TierStarts;
                AddAllMultiDependentDataKeys(allData, consumerClassRateStructures, fnGetList, consumerClass, "tier_starts");

                fnGetList = rs => rs.TierStartsCommodity;
                AddAllMultiDependentDataKeys(allData, consumerClassRateStructures, fnGetList, consumerClass, "tier_starts_commodity");

                fnGetList = rs => rs.TierStartsDrought;
                AddAllMultiDependentDataKeys(allData, consumerClassRateStructures, fnGetList, consumerClass, "tier_starts_drought");

                fnGetList = rs => rs.UtilitySurcharge;
                AddAllMultiDependentDataKeys(allData, consumerClassRateStructures, fnGetList, consumerClass, "utility_surcharge");
            }

            // row names
            worksheet.Cells[1, 1].Value = "consumer class";
            worksheet.Cells[2, 1].Value = "property";
            worksheet.Cells[3, 1].Value = "depends on key";
            worksheet.Cells[4, 1].Value = "depends on value";
            for (int row = 0; row < totalRows; row++)
            {
                worksheet.Cells[5 + row, 1].Value = agencyObjects[row].Metadata.UtilityName;
            }

            // row names
            int index = 0;
            foreach (var kvp in allData)
            {
                worksheet.Cells[1, 2 + index].Value = kvp.Key.Item1; // consumer class
                worksheet.Cells[2, 2 + index].Value = kvp.Key.Item2; // property
                worksheet.Cells[3, 2 + index].Value = kvp.Key.Item3; // depends on key
                worksheet.Cells[4, 2 + index].Value = kvp.Key.Item4; // depends on value
                index++;
            }

            // cell values
            index = 0;
            foreach (var kvp in allData)
            {
                for (int row = 0; row < totalRows; row++)
                {
                    worksheet.Cells[5 + row, 2 + index].Value = kvp.Value(agencyObjects[row]);
                }
                index++;
            }
        }

        private static void AddAllMultiDependentDataKeys(Dictionary<Tuple<string, string, string, string>, Func<AgencyObject, string>> allData,
            List<AgencyRateStructure> consumerClassRateStructures, Func<AgencyRateStructure, AgencyList> fnGetList, string consumerClass, string propertyName)
        {
            if (consumerClassRateStructures.Any(rs => fnGetList(rs) != null))
            {
                List<string> allDependsOn = GetAllDependsOn(consumerClassRateStructures, fnGetList);
                foreach (string dependsOn in allDependsOn)
                {
                    List<string> keys = GetDependsOnKeys(consumerClassRateStructures, fnGetList, dependsOn);
                    foreach (string key in keys)
                        allData.Add(Tuple.Create(consumerClass, propertyName, dependsOn, string.IsNullOrEmpty(dependsOn) && key == "All" ? "" : key), obj => GetCommaDelimitedListValue(consumerClass, obj, rs => fnGetList(rs), key, dependsOn));
                }
            }
        }

        private static List<string> GetAllDependsOn(List<AgencyRateStructure> consumerClassRateStructures, Func<AgencyRateStructure, AgencyList> fnGetList)
        {
            return consumerClassRateStructures
                .Where(rs => fnGetList(rs) != null)
                .Select(rs => string.Join("|", (fnGetList(rs) ?? new AgencyList()).DependsOn ?? new List<string>()))
                .Distinct().ToList();
        }

        private static List<string> GetDependsOnKeys(List<AgencyRateStructure> consumerClassRateStructures, Func<AgencyRateStructure, AgencyList> fnGetList, string dependsOn)
        {
            return consumerClassRateStructures
                .Where(rs => fnGetList(rs) != null && dependsOn == string.Join("|", (fnGetList(rs).DependsOn ?? new List<string>())))
                .SelectMany(rs => fnGetList(rs).Values.Keys)
                .Distinct().ToList();
        }

        private static List<string> GetDependentKeys(IEnumerable<AgencyRateStructure> allRateStructures, Func<AgencyRateStructure, AgencyList> fnGetList)
        {
            return allRateStructures.Select(x => fnGetList(x)).Where(x => x != null).SelectMany(x => x.Values.Keys).Distinct().ToList();
        }

        private static string GetCommaDelimitedListValue(string consumerClass, AgencyObject obj, Func<AgencyRateStructure, AgencyList> fnGetList, string key, string dependsOn = null)
        {
            AgencyRateStructure rateStructure = GetRateStructure(obj, consumerClass);
            return rateStructure != null
                ? string.Join(", ", GetValueList(fnGetList(rateStructure), key, dependsOn))
                : null;
        }

        private static string GetFirstListValue(string consumerClass, AgencyObject obj, Func<AgencyRateStructure, AgencyList> fnGetList, string key, string dependsOn = null)
        {
            AgencyRateStructure rateStructure = GetRateStructure(obj, consumerClass);
            return rateStructure != null
                ? GetValueList(fnGetList(rateStructure), key, dependsOn).FirstOrDefault()
                : null;
        }

        private static List<string> GetValueList(AgencyList list, string key, string dependsOn = null)
        {
            if (string.IsNullOrWhiteSpace(dependsOn))
                dependsOn = null;

            if (dependsOn != null && (list == null || list.DependsOn == null || dependsOn != string.Join("|", list.DependsOn)))
                return new List<string>();

            return list != null && list.Values != null && list.Values.ContainsKey(key) ? list.Values[key] : new List<string>();
        }

        private static AgencyRateStructure GetRateStructure(AgencyObject obj, string consumerClass)
        {
            return obj != null ? obj.RateStructures.FirstOrDefault(x => x.ConsumerClass == consumerClass) : null;
        }

        private static List<string> GetListKeys(IEnumerable<AgencyList> lists)
        {
            return lists.Where(x => x != null).SelectMany(x => x.Values.Keys).Distinct().ToList();
        }
    }
}
