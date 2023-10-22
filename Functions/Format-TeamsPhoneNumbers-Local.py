import phonenumbers
import json
import subprocess

def getGitRoot():
    return subprocess.Popen(['git', 'rev-parse', '--show-toplevel'], stdout=subprocess.PIPE).communicate()[0].rstrip().decode('utf-8')

git_root = getGitRoot()

# Read phone numbers from a TXT file
with open(f"{git_root}/.local/TeamsPhoneNumberOverview_AllCsOnlineNumbers.txt", "r") as txt_file:
    phone_numbers_str = txt_file.read().strip()
    phone_numbers_list = phone_numbers_str.split(';')

# Create a list to store dictionaries containing original and formatted numbers
output_list = []

# Iterate through each phone number and process/format it
for index, phone_number in enumerate(phone_numbers_list, start=1):
    try:
        parsed_number = phonenumbers.parse(phone_number, None)
        formatted_number = phonenumbers.format_number(parsed_number, phonenumbers.PhoneNumberFormat.INTERNATIONAL)
        formatted_number = formatted_number.replace('-', ' ')
        entry = {"original": phone_number, "formatted": formatted_number}
        output_list.append(entry)

        # Print progress every 100 elements
        if index % 100 == 0:
            print(f"Processed {index} out of {len(phone_numbers_list)} elements")
    except phonenumbers.NumberFormatException:
        print(f"Invalid phone number: {phone_number}")

# Serialize the list to a JSON string
json_content = json.dumps(output_list, indent=4)

# Write the JSON content enclosed in single quotes to a text file
with open(f"{git_root}/.local/TeamsPhoneNumberOverview_PrettyNumbers.txt", "w") as output_txt_file:
    output_txt_file.write(f"'{json_content}'")