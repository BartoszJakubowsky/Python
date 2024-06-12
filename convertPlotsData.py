import re

def parse_file(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    parcels = []
    #single line data
    new_parcel_pattern = re.compile(r"Informacja z opisowych danych ewidencji gruntów i budynków udostępniona jednostce wykonawstwa geodezyjnego w")
    unit_name_pattern = re.compile(r"Jednostka ewidencyjna : (.+)")
    district_name_pattern = re.compile(r"Nazwa obrębu : (.+)")
    district_number_pattern = re.compile(r"Numer obrębu : (.+)")

    #multiline data
    plot_owner_number = re.compile(r"^\d+$")
    country_plot_owner = re.compile(r"\d+\s+skarb państwa", re.IGNORECASE)
    pesel_pattern = re.compile(r"PESEL: (\d+)")
    nip_pattern = re.compile(r"NIP: (\d+)")
    owner_section_pattern = re.compile('Udział')

    default_parcel = {
                'unit_name': '',
                'district_name': '',
                'district_number': '',
                'plot_name': '',
                'owners': []
            }

    global current_parcel
    global multi_line_data
    global current_plot_owners

    def reset_current_parcel(line):

        if new_parcel_pattern.search(line) and len(parcels) != 0:
                    parcels.append(current_parcel)
                    current_parcel = default_parcel
                    is_plot_owners_active = False
                    current_plot_owners = []
                    
    def extract_single_line_data(line):
        unit_name = unit_name_pattern.search(line)
        district_name = district_name_pattern.search(line)
        district_number = district_number_pattern.search(line)
        pesel = pesel_pattern.search(line)
        nip = nip_pattern.search(line)
        owner_section = owner_section_pattern.search(line)

        if unit_name:
            default_parcel['unit_name'] = unit_name.group(1)
        if district_name:
            default_parcel['district_name'] = district_name.group(1)
        if district_number:
            default_parcel['district_number'] = district_number.group(1)

    # def extract_multi_line_data(line):
         

    def check_multiline_data(line):
         if owner_section_pattern.search(line)
            return True

        
        # if plot_owner_number.search(line) or country_plot_owner.search(line)
    for line in lines:
        reset_current_parcel(line)

        if check_multiline_data(line): multi_line_data = True

        # if multi_line_data:
            # extract_multi_line_data(line)
        # else:
        extract_single_line_data(line)
       

file_path = "E:\\Downloads\\test.txt"
parcels = parse_file(file_path)

