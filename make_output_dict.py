def make_output_dict(Outputs_config, time_series_keys):
    output_dict = {}
    for output_keys in Outputs_config:
        if type in (alasdf, asdlf):
            time_to_val = []
        for timeval in time_series_keys:
            time_to_val[timeval] = Outputs_config[output_keys]['Value']
        time_to_val.append(Outputs_config[output_keys]['Given_values'])
        output_dict[(Outputs_config[output_keys]['Display_name'])] = time_to_val
        if Outputs_config[output_dict]['Calc_method'] == 'Check_system_find_year':
            input2keyname = input_dicts[Outputs_config][output_dict]['Input1_table_name']['Input1_key_name']
            yearvaluedict = input_dicts[Outputs_config][output_dict]['Input2_table_name'][input2keyname]
            output_values_dict = []
            for timekey in timeseriesdict:
                output_values_dict[timekey] = yearvaluedict[timeseriesdict[timekey]["Year"]]
            output_values_dict.append(Outputs_config[output_keys]['Given_values'])
            Outputs_config[output_dict]['Value'] = output_values_dict


