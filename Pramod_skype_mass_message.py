from skpy import Skype

# Skype credentials
username = "lingaraj.sahu_tal@outlook.com"  # Replace with your Skype username
password = "Sahu@321"  # Replace with your Skype password

# Employees (Skype usernames) to send messages to
employees = ["jaydeep.parakh_tal", "live:indu.singh_tal", "live:Muktambree.Juneja_tal", "live:prasad.puro_tal", "live:.cid.d2db34a3148abf9e", "live:.cid.788e0f53da6c8bc6", "live:.cid.d7247f0ad4399955", "live:achyut.gosavi_tal", "live:.cid.77831c3f087915f7", "live:.cid.75dce8cefca30652", "live:.cid.de75e2342b2eab1d", "live:.cid.aa035431d0b3bf4d", "live:.cid.e181872da7ead3ea", "live:.cid.c179f61db24778ee", "live:.cid.a5688a0676c0e983", "live:.cid.a0d5485623467960", "live:.cid.d71b9dd25d444864", "live:.cid.793aa03a7627456d", "live:.cid.eac7bfe300da7fe2", "live:.cid.c18198cbc96eed94", "live:.cid.b7f564fbb6a25ec1", "live:.cid.6de2fd9b072b745d", "live:.cid.f466cec02f0c57bd", "live:.cid.a478fe6546002185", "live:.cid.3621be03f3e3e4cc", "live:.cid.81a04ce19a69eb31", "live:.cid.4f6d763ace60f940", "live:.cid.3fb6860b6b095772", "live:.cid.dcefc209c0a4d1af", "live:.cid.311dd8be5601cc84", "live:.cid.6e99677db392db34", "poushali.ganguly_tal", "namrata.gupta_tal", "live:.cid.259f0d5493fad52f", "live:.cid.c1b66aeae4cc8460", "live:.cid.6600bef41e16727f", "live:.cid.fcb901f03213a37f"]


# Message to send
message = "Gentle reminder..!\n\n Need your help to complete Physical verification of your Talentica provided Laptop/s. Please fill out the form, which need max 15 seconds. Instructions are given for your reference.\nPlease add 2 pictures\n1.  Front while it is turned on\n2. Bottom/Rear side where the IT Asset sticker is attached.\n\nNote: CPU: , Asset Tag starts with Ux/CP/xx-xx/ \n\nIn case you do have multiple assets please repeat the same process.\n https://forms.office.com/r/vNVQzKy3WF \n\nNOTE: If you do not use  Talentica provided laptop and receives this message please ping us from skype.\nPlease ignore if you already filled this form.\n\nThank you!"

# Log into Skype
try:
    skype = Skype(username, password, tokenFile="skype_token.txt")  # Login and save token for future sessions
except Exception as e:
    print(f"Failed to log into Skype: {e}")
    exit(1)

# Send message to each employee
for employee in employees:
    try:
        # Check if the contact exists
        contact = next((c for c in skype.contacts if c.id == employee), None)
        if contact:
            # If the contact exists, send a message
            chat = contact.chat
            chat.sendMsg(message)
            print(f"Message sent to {employee}")
        else:
            print(f"Contact {employee} is not in your contact list. Please add them first.")
    except Exception as e:
        print(f"Failed to send message to {employee}: {e}")

print("Finished sending messages.")
