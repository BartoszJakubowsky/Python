import re

def parse_file(file_path):
    parcels = []
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()
    
    parcel_pattern = re.compile(r"Informacja z opisowych danych ewidencji gruntów i budynków udostępniona jednostce wykonawstwa geodezyjnego w")
    unit_pattern = re.compile(r"Jednostka ewidencyjna : (.+)")
    district_name_pattern = re.compile(r"Nazwa obrębu : (.+)")
    district_number_pattern = re.compile(r"Numer obrębu : (.+)")
    owner_start_pattern = re.compile(r"(\d+)")
    pesel_pattern = re.compile(r"PESEL: (\d+)")
    
    current_parcel = None
    owners = []
    capture_owners = False

    for line in lines:
        if parcel_pattern.search(line):
            if current_parcel:
                parcels.append(current_parcel)
            current_parcel = {
                'unit': '',
                'district_name': '',
                'district_number': '',
                'owners': []
            }
            owners = []
            capture_owners = False

        if unit_pattern.search(line):
            current_parcel['unit'] = unit_pattern.search(line).group(1)

        if district_name_pattern.search(line):
            current_parcel['district_name'] = district_name_pattern.search(line).group(1)

        if district_number_pattern.search(line):
            current_parcel['district_number'] = district_number_pattern.search(line).group(1)
        
        if owner_start_pattern.match(line.strip()) and not capture_owners:
            capture_owners = True
            continue
        
        if capture_owners:
            match = pesel_pattern.search(line)
            if match:
                owners.append({
                    'owner_info': '',
                    'parents': '',
                    'address': '',
                    'PESEL': match.group(1),
                    'ownership': ''
                })
            elif 'PESEL' in line:
                current_owner = owners[-1]
                current_owner['owner_info'] = line.strip()
                address_line = lines[lines.index(line) + 1].strip()
                current_owner['address'] = address_line
                ownership_line = lines[lines.index(line) + 2].strip()
                current_owner['ownership'] = ownership_line

            if not owner_start_pattern.match(line.strip()) and capture_owners and not pesel_pattern.search(line):
                capture_owners = False
                current_parcel['owners'] = owners

    if current_parcel:
        parcels.append(current_parcel)
    
    return parcels

# Usage example
file_path = "C:\\Users\\BFS\Downloads\\test.txt"
parcels = parse_file(file_path)
for parcel in parcels:
    print(parcel)
