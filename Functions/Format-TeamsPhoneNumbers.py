import phonenumbers
import json
import automationassets

# Get the input phone numbers string and split it into a list
phone_numbers_str = automationassets.get_automation_variable("TeamsPhoneNumberOverview_AllCsOnlineNumbers")
phone_numbers_list = phone_numbers_str.split(';')

# Create a list to store dictionaries containing original and formatted numbers
output_list = []

# Iterate through each phone number and process/format it
for phone_number in phone_numbers_list:
    try:
        parsed_number = phonenumbers.parse(phone_number, None)
        formatted_number = phonenumbers.format_number(parsed_number, phonenumbers.PhoneNumberFormat.INTERNATIONAL)
        formatted_number = formatted_number.replace('-', ' ')
        entry = {"original": phone_number, "formatted": formatted_number}
        output_list.append(entry)
    except phonenumbers.NumberFormatException:
        print(f"Invalid phone number: {phone_number}")

# Serialize the list to a JSON string
json_content = json.dumps(output_list, indent=4)

# Set the JSON content as the value of the Azure Automation variable
automationassets.set_automation_variable("TeamsPhoneNumberOverview_PrettyNumbers", f"'{json_content}'")