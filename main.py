import gspread
from oauth2client.service_account import ServiceAccountCredentials
import yaml
from yaml import SafeLoader


# Read data from a Google Sheet and return a list of dictionaries
def read_google_sheet(sheet_name):
    scope = ['https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive.file",
             "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name('anny.json', scope)
    client = gspread.authorize(creds)
    sheet = client.open(sheet_name).sheet1
    records = sheet.get_all_records()
    return records

# Create a CloudFormation template in JSON format from the given data
def create_cfn_template(vpc, public_subnet, private_subnet, route_table, security_group):
    heading = {'AWSTemplateFormatVersion': '2010-09-09'}

    nested_json = {'Resources': {}}
    name_vpc = vpc[0]['Name']
    name_public_subnet = public_subnet[0]['Name']
    name_private_subnet = private_subnet[0]['Name']
    name_route_table = route_table[0]['Name']
    name_security_group = security_group[0]['Name']

    vpc[0].pop("Name")
    public_subnet[0].pop("Name")
    private_subnet[0].pop("Name")
    route_table[0].pop("Name")
    security_group[0].pop("Name")

    for item in vpc:
        for k,v in item.items():
            if k == 'EnableDnsHostnames' and v == 'Yes':
                item = {
                    **item,
                    'EnableDnsHostnames': True
                }
                # print(item)
            if k == 'EnableDnsSupport' and v == 'Yes':
                item = {
                    **item,
                    'EnableDnsSupport': True
                }
                # print(item)

    vpc[0] = item

    nested_json['Resources'][name_vpc] = {'Type': 'AWS::EC2::VPC', 'Properties': vpc[0]}
    nested_json['Resources'][name_public_subnet] = {'Type': 'AWS::EC2::Subnet','Properties': public_subnet[0]}
    nested_json['Resources'][name_private_subnet] = {'Type': 'AWS::EC2::Subnet','Properties': private_subnet[0]}
    nested_json['Resources'][name_route_table] = {'Type': 'AWS::EC2::RouteTable','Properties': route_table[0]}
    nested_json['Resources'][name_security_group] = {'Type': 'AWS::EC2::SecurityGroup','Properties': security_group[0]}

    nested_json['Resources'][name_public_subnet]['Properties'].update({'VpcId': { "Ref": str(name_vpc)}})
    nested_json['Resources'][name_private_subnet]['Properties'].update({'VpcId': { "Ref": str(name_vpc)}})
    nested_json['Resources'][name_route_table]['Properties'].update({'VpcId': { "Ref": str(name_vpc)}})
    nested_json['Resources'][name_security_group]['Properties'].update({'VpcId': {"Ref": str(name_vpc)}})

    pub_subnet = list(yaml.load_all(public_subnet[0]['Tags'], Loader=SafeLoader))[0]
    public_subnet[0]['Tags'] = pub_subnet

    puv_subnet = list(yaml.load_all(private_subnet[0]['Tags'], Loader=SafeLoader))[0]
    private_subnet[0]['Tags'] = puv_subnet

    route = list(yaml.load_all(route_table[0]['Tags'], Loader=SafeLoader))[0]
    route_table[0]['Tags'] = route

    sg = list(yaml.load_all(security_group[0]['SecurityGroupIngress'], Loader=SafeLoader))[0]

    security_group[0]['SecurityGroupIngress'] = sg

    sg1 = list(yaml.load_all(security_group[0]['SecurityGroupEgress'], Loader=SafeLoader))[0]
    security_group[0]['SecurityGroupEgress'] = sg1

    heading.update(nested_json)
    return heading


# Write a CloudFormation template to a YAML file
def write_yaml_file(cfn_template, file_name):
    with open(file_name, 'w') as f:
        yaml.dump(cfn_template, f, default_flow_style=False, explicit_end=False,sort_keys=False)

# Read data from Google Sheets
vpc = read_google_sheet('VPC_creation_anant')
public_subnet = read_google_sheet('Public_Subnet')
private_subnet = read_google_sheet('Private_Subnet')
route_table = read_google_sheet('Route_Table')
security_group = read_google_sheet('Security_group')

# Create CloudFormation template from data
cfn_template = create_cfn_template(vpc, public_subnet, private_subnet, route_table, security_group)

# Write CloudFormation template to YAML file
write_yaml_file(cfn_template, 'cloudformation_template.yaml')
