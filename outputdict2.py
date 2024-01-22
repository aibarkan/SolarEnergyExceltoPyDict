<<<<<<< HEAD

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
=======

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
>>>>>>> 9d4deffa7a31fe233681fa0d7fff9a4605386c0e
