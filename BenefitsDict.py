import openpyxl
import InputDicts
from copy import deepcopy
import json

#CONSTANTS
Quarters_range = (2018, 2022)
Years_range = (2022, 2030)
Given_Table_Col = ['Units','Source']
val = 'Value'


Outputs_config = {
    "Change_Peak_Load": {
        'Calc_method': "Single_input_val",
        'Input_table_name': 'Technology Assumptions',
        'Inputdata': 'Inputdata',
        'Value_name': 'Change in Peak Load (Effective Project Impact on Substation Peak)',
        "Display_name": 'Change Peak Load (Y, r)',
        'Given_values': {"Units": "DMW", "Source": "Inputs - Project Specific"},
        'Value': None},
    "1_Loss": {
        'Calc_method': "Single_input_val_and_calc",
        'Outer_input_name': 'Inputs - Fixed',
        'Inputdata': 'Inputdata',
        'Input_table_name': 'General Assumptions',
        'Value_name': 'Loss Percentage (based on location of NWA)',
        'Display_name': '1-Loss % (Y, b -> r)',
        'Given_values': {"Units": "DMW", "Source": "Inputs - Project Specific"},
        'Value': None},
    "Distribution_Confidence_Factor": {
        'Calc_method': "Single_input_val",
        'Outer_input_name': 'Inputs - Project Specific',
        'Inputdata' : 'Inputdata',
        'Input_table_name': 'Project-Specific Assumptions',
        'Value_name': 'Distribution Coincidence Factor',
        'Display_name': 'Distribution Coincidence Factor (C, V, Y)',
        'Given_values': {"Units": "DMW", "Source": "Inputs - Project Specific"},
        'Value': None},
    "Derating_Factor": {
        'Calc_method': "Single_input_val",
        'Outer_input_name': 'Inputs - Project Specific',
        'Input_table_name': 'Project-Specific Assumptions',
        'Inputdata' : 'Inputdata',
        'Value_name': 'Derating Factor',
        'Display_name': 'Derating Factor (Z, Y)',
        'Given_values': {"Units": "DMW", "Source": "Inputs - Project Specific"},
        'Value': None},
    "Marginal_Distribution_Factor": {
        'Calc_method': "Check_system_find_year",
        'Input1_table_name': 'Technology Assumptions',
        'Input1_key_name': 'Retail delivery (NWA connected to:)',
        'Input2_table_name': 'Utility System Average Marginal Cost of Service ($/kW-yr)',
        'Inputdata': 'Inputdata',
        'Display_name': 'Marginal Distribution Cost (C, V, Y, b)',
        'Given_values': {"Units": "DMW", "Source": "Inputs - Project Specific"},
        'Value': None},
    "1/kw_constant": {
        'Calc_method': "Constant",
        'Display_name': '1/kW -> 1/MW',
        'Input_table_name': 'Project-Specific Assumptions',
        'Given_values': {"Units": "DMW", "Source": "Inputs - Project Specific"},
        'Value': 1000},
    "Quarterly_Benefit_Factor": {
        'Calc_method': "Constant",
        'Display_name': '1/kW -> 1/MW',
        'Input_table_name': 'Project-Specific Assumptions',
        'Given_values': {"Units": "DMW", "Source": "Inputs - Project Specific"},
        'Value': 1000},

}

def main():
    input_dicts = InputDicts.main()
    timeserieskeys = get_time_series_keys()
    Outputs_dict = populate_output_config_dict(input_dicts, timeserieskeys)
    #print(Outputs_config)
    output_dict = make_output_dict(Outputs_config, timeserieskeys)
    #print(output_dict)
    json_object = json.dumps(output_dict, indent = 4)
    #print(json_object)



def populate_output_config_dict(input_dicts, timeseriesdict):
    for output_dict in Outputs_config.keys():
        if Outputs_config[output_dict]['Calc_method'] == 'Check_system_find_year':
            input2keyname = input_dicts[Outputs_config][output_dict]['Input1_table_name']['Input1_key_name']
            yearvaluedict = input_dicts[Outputs_config][output_dict]['Input2_table_name'][input2keyname]
            output_values_dict = []
            for timekey in timeseriesdict:
                output_values_dict[timekey] = yearvaluedict[timeseriesdict[timekey]["Year"]]
            output_values_dict.append(Outputs_config[output_keys]['Given_values'])
            Outputs_config[output_dict]['Value'] = output_values_dict


            innerdict_tablenamekey1 = Outputs_config[output_dict]['Input_table_name1']
            innerdict_inputnamekey = Outputs_config[output_dict]['Inputdata']
            innerdict_valuecolumnname1 = Outputs_config[output_dict]['Value_name1']
            innerdict_tablenamekey2 = Outputs_config[output_dict]['Input_table_name2']
            innerdict_valuecolumnname2 = Outputs_config[output_dict]['Value_name2']
            innerdict_valuecolumnname3 = Outputs_config[output_dict]['Value_name3']
            if input_dicts[innerdict_tablenamekey1][innerdict_inputnamekey][innerdict_valuecolumnname1][val] == innerdict_valuecolumnname2:
                for year in input_dicts[innerdict_tablenamekey2][innerdict_inputnamekey][innerdict_valuecolumnname2]:
                    for quarter_year in range(Quarters_range[0], Quarters_range[1]):
                        if quarter_year == year:
                            for quarter in range(1, 5):
                                Outputs_config[output_dict]['Value'] = input_dicts[innerdict_tablenamekey2][innerdict_inputnamekey][innerdict_valuecolumnname2][year]
                                print(Outputs_config[output_dict]['Value'])
                                #print(input_dicts[innerdict_tablenamekey2][innerdict_inputnamekey][innerdict_valuecolumnname2][year])
                    for year_val in range(Years_range[0],Years_range[1]):
                        if year_val == year:
                            Outputs_config[output_dict]['Value'] = input_dicts[innerdict_tablenamekey2][innerdict_inputnamekey][innerdict_valuecolumnname2][year]
            else:
                for year in input_dicts[innerdict_tablenamekey2][innerdict_inputnamekey][innerdict_valuecolumnname3]:
                    for quarter_year in range(Quarters_range[0], Quarters_range[1]):
                        if quarter_year == year:
                            for quarter in range(1, 5):
                                Outputs_config[output_dict]['Value'] = input_dicts[innerdict_tablenamekey2][innerdict_inputnamekey][innerdict_valuecolumnname2][year]
                    for year_val in range(Years_range[0],Years_range[1]):
                        if year_val == year:
                            Outputs_config[output_dict]['Value'] = input_dicts[innerdict_tablenamekey2][innerdict_inputnamekey][innerdict_valuecolumnname2][year]
        elif Outputs_config[output_dict]['Calc_method'] != 'Constant':
            innerdict_tablenamekey = Outputs_config[output_dict]['Input_table_name']
            innerdict_inputnamekey = Outputs_config[output_dict]['Inputdata']
            innerdict_valuecolumnname = Outputs_config[output_dict]['Value_name']
            if Outputs_config[output_dict]['Calc_method'] == 'Single_input_val':
                Outputs_config[output_dict]['Value'] = input_dicts[innerdict_tablenamekey][innerdict_inputnamekey][innerdict_valuecolumnname]['Value']
            elif Outputs_config[output_dict]['Calc_method'] == 'Single_input_val_and_calc':
                Outputs_config[output_dict]['Value'] = "{:.2%}".format((1-(float(input_dicts[innerdict_tablenamekey][innerdict_inputnamekey][innerdict_valuecolumnname]['Value']))))
            else:
                print('nothing yet')
        else:
            print('nothing yet')
    return Outputs_config

def get_time_series_keys():
    time_series_keys={}
    for quarter_year in range(Quarters_range[0], Quarters_range[1]):
        for quarter in range(1, 5):
            time_series_keys[str(quarter_year) + "QTR" + str(quarter)]={"Year": quarter_year, "Quarter": "QTR" + str(quarter), "TimeGranularity": "Quarter"}
    for year in range(Years_range[0],Years_range[1]+1):
        time_series_keys[str(year)] = {"Year": year, "TimeGranularity": "Year"}
    return time_series_keys


def make_output_dict(Outputs_config, time_series_keys):
    output_dict = {}
    for output_keys in Outputs_config:
        if type in 'single value':
            output_values_dict[time_series_keys.keys()] = Outputs_config[output_keys]['Value']
            #for timeval in time_series_keys.keys():
            #    output_values_dict[timeval] = Outputs_config[output_keys]['Value']
            output_values_dict.append(Outputs_config[output_keys]['Given_values'])
            output_dict[(Outputs_config[output_keys]['Display_name'])] = output_values_dict

        if Outputs_config[output_dict]['Calc_method'] == 'Check_system_find_year':
            input2keyname = input_dicts[Outputs_config][output_dict]['Input1_table_name']['Input1_key_name']
            yearvaluedict = input_dicts[Outputs_config][output_dict]['Input2_table_name'][input2keyname]
            output_values_dict = []
            for timekey in timeseriesdict:
                output_values_dict[timekey] = yearvaluedict[timeseriesdict[timekey]["Year"]]
            output_values_dict.append(Outputs_config[output_keys]['Given_values'])
            Outputs_config[output_dict]['Value'] = output_values_dict
    return output_dict


def OLDmake_output_dict(Outputs_config, time_series_keys):
    output_dict = {}
    for output_keys in Outputs_config:
        time_to_val = []
        for timeval in time_series_keys:
            time_to_val[timeval] = Outputs_config[output_keys]['Value']
            time_to_val.append(deepcopy(val_dict))
        output_dict[(Outputs_config[output_keys]['Display_name'])] = {'Values': time_to_val,'Given Values': Outputs_config[output_keys]['Given_values']}

    #output_dict['Total Nominal Benefits'] = sum('totalbenefit')
    #output_dict['Total NPV Benefit'] = npv()
    return output_dict



main()
