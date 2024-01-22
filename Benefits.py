import openpyxl
import TableName


#CONSTANTS
Quarters_range = (2018, 2022)
Years_range = (2023, 2030)
Given_Table_Col = ['Units','Source']

Outputs_config = {
    "Change_Peak_Load" : {
        'Calc_method': "Single_input_val",
        'Outer_input_name': 'Inputs - Project Specific',
        'Input_table_name': 'Technology Assumptions',
        'Value_name': 'Change in Peak Load (Effective Project Impact on Substation Peak)',
        'Display_name': 'Change Peak Load (Y, r)',
        'Given_values': {"Units": "DMW", "Source": "Inputs - Project Specific"},
        'Value': None},
    "1_Loss" : {
        'Calc_method': "Single_input_val",
        'Outer_input_name': 'Inputs - Fixed',
        'Input_table_name': 'General Assumptions',
        'Value_name': 'Loss Percentage (based on location of NWA)',
        'Display_name': '1-Loss % (Y, b -> r)',
        'Given_values': {"Units": "DMW", "Source": "Inputs - Project Specific"},
        'Value': None},
    "Distribution_Confidence_Factor" : {
        'Calc_method': "Single_input_val",
        'Outer_input_name': 'Inputs - Project Specific',
        'Input_table_name': 'Project Specific Assumptions',
        'Value_name': 'Distribution Coincidence Factor',
        'Display_name': 'Distribution Coincidence Factor (C, V, Y)',
        'Given_values': {"Units": "DMW", "Source": "Inputs - Project Specific"},
        'Value': None},
    "Derating_Factor" : {
        'Calc_method': "Single_input_val",
        'Outer_input_name': 'Inputs - Project Specific',
        'Input_table_name': 'Project Specific Assumptions',
        'Value_name': 'Derating Factor',
        'Display_name': 'Derating Factor (Z, Y)',
        'Given_values': {"Units": "DMW", "Source": "Inputs - Project Specific"},
        'Value': None},
    "Marginal_Distribution_Factor" : {
        'Calc_method': "tricky",
        'Outer_input_name': 'Inputs - Project Specific',
        'Input_table_name': 'Project Specific Assumptions',
        'Value_name': 'Derating Factor',
        'Display_name': 'Marginal Distribution Cost (C, V, Y, b)',
        'Given_values': {"Units": "DMW", "Source": "Inputs - Project Specific"},
        'Value': None},
    "1/kw_constant" : {
        'Calc_method': "Single_input_val",
        'Display_name': '1/kW -> 1/MW',
        'Given_values': {"Units": "DMW", "Source": "Inputs - Project Specific"},
        'Value': None}

}

def main():
    #print(TableName.get_job_configuration_values())

def populate_output_config_dict(input_dicts):
    for val_dict in Outputs_config:
        if val_dict['Calc_method'] == "Single_input_val":
            print(input_dicts[0])
            print(input_dicts['Outer_input_name']['Input_table_name']['Value_name'])
            #val_dict['Value'] = input_dict['Outer_input_name'][][]
        else:
            print('Complicated func')

#def make_date_dict(Quarters_range, Years_range):
    #date_dict = {}
    #for x in range



