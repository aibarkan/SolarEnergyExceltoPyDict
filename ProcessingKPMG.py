import openpyxl
import Input



def main():
    input_dicts = Input.main()
    Outputs_dict = populate_output_config_dict(input_dicts)
    #account name-houeholdname-[billing zip/postal code]

def populate_output_config_dict(input_dicts):
    '''
    outerdictkey = input_dicts['AccountName']
    for x in input_dicts
        outerdictkey = input_dicts['AccountName']
        innerdictkey = input_dicts[outerdictkey][]
    innerdictkey = input_dicts['AccountName']
    '''
    innerlist = []
    for x in input_dicts['Account Name'].values():
        for y in x:
            innerlist.append(y)
    for x in input_dicts['Account Name']:
        for z in innerlist:
            print(z)

