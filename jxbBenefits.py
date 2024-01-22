QuartersRange = 2018,2022
YearsRange = 2023,2030
NonComputedTableColumns=['Units','Source','Val']
TotalsOnlyColumns=['Total Nominal Benefits','Total NPV Benefit']

OutputsConfiguration={
    "ChangePeakLoad":{
        "SortOrder":1,
        "TotalBenefitCalcMethod": "Numerator",
        "TimeSeriesCalcMethod": "SingleInputValue",
        "TimeSeriesInputValueName": "Change in Peak Load",
        "TimeSeriesSourceInputTableName": "Technology Assumptions",
        "Display Label": "Change Peak Load (Y, r)",
        "NonComputedValues": {"Units": "DMW", "Source": "Inputs - Project Specific"}},
    "LossPercentage":{
        "SortOrder":2,
        "TotalBenefitCalcMethod": "DenominatorProduct",
        "TimeSeriesCalcMethod": "SingleInputValue",
        "TimeSeriesSourceInputValueName": "Loss Percentage",
        "TimeSeriesSourceInputTableName": "General Assumptions",
        "Display Label": "1-Loss % (Y, b -> r)",
        "NonComputedValues": {"Units": "%", "Source": "BCA Handbook"}},
    "MarginalDistributionCost":{
        "SortOrder":5,
        "TotalBenefitCalcMethod": "DenominatorProduct",
        "TimeSeriesCalcMethod": "InputValueWithInputDependencyAndYearlyValues",
        "TimeSeriesSourceInputValueName": "Retail delivery",
        "TimeSeriesSourceInputTableName": "Technology Assumptions",
        "TimeSeriesDependentInputTableName": "General Assumptions",
        "Display Label": "Marginal Distribution Cost (C, V, Y, b)",
        "NonComputedValues": {"Units": "$/kW-yr", "Source": "Inputs - Fixed"}},
    "1kw-1mw":{
        "SortOrder":2,
        "TotalBenefitCalcMethod": "DenominatorProduct",
        "TimeSeriesCalcMethod": "Constant",
        "TimeSeriesConstantValue": 1000,
        "Display Label": "1/kW -> 1/MW",
        "NonComputedValues": {"Units": "KW/MW", "Source": "Constant"}}
    }


def get_time_series_keys():
    timeserieskeys={}
    for quarteryear in range(QuartersRange[0],QuartersRange[1]):
        for quarter in range(1, 4):
            timeserieskeys[str(quarteryear) + "QTR" + str(quarter)]={"Year":quarteryear, "Quarter": "QTR"+str(quarter), "TimeGranularity":"Quarter"}
    for year in range(YearsRange[0],YearsRange[1]):
        timeserieskeys[str(year)]={"Year":year, "TimeGranularity":"Year"}
    return timeserieskeys

def main():
    #print(OutputsConfiguration)
    timekeys=get_time_series_keys()
    print(timekeys)

main()
