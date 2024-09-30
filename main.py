from datetime import datetime
from sendLetterAttachedMail import sendPdfAttachmentMail
from createLetter import getLetter
from csvWriter import writeRecordsToCsv, readRecordsFromCsv, readRecordsFromExcel
from emailHtmlContent import getHtmlContent
import easygui

whatsApplinks = {
    "Capture The Flag": "https://chat.whatsapp.com/HMZjS6mow0L1iLMEuLdCIu",
    "App Development Competition": "https://chat.whatsapp.com/KmH0VSqOlmIDYgYdQ55ew1",
    "Code Sprint": "https://chat.whatsapp.com/K3c5TY3U7wp2AZsQBfmu75",
    "Code in the Dark": "https://chat.whatsapp.com/LNz8GR6rdMa60QYwLnn8qf",
    "Data Science": "https://chat.whatsapp.com/EsmHB6N9lhuJ191i2fIvts",
    "Data Visualisation": "https://chat.whatsapp.com/EPyXqCpjdIREVJJlOOkVr5",
    "Database Design": "https://chat.whatsapp.com/D2hzoFOO8KN0KbGI0KRks0",
    "SyncOS": "https://chat.whatsapp.com/Exn0o0SzRIZ22vJQRfnDhS",
    "EA": "https://chat.whatsapp.com/LIGK5fhE8GC07ly0tjrJAL",
    "EM Career Fair": "https://chat.whatsapp.com/EJwf4yGFS2mBX8MWzjdtSm",
    "EM Competition": "https://chat.whatsapp.com/LGGnRaknt8h3cI5CfDeESN",
    "FYP Exhibition": "https://chat.whatsapp.com/HHRdPM6py9LCteAhPFD6tn",
    "Food": "https://chat.whatsapp.com/HVBs9xHjlMCFChGQpVp8pS",
    "General Competitions": "https://chat.whatsapp.com/GWgukNPfxMc2Ao9YzrKBlN",
    "Creativity": "https://chat.whatsapp.com/BDeErfN10ii5y1nTzRq2IL",
    "Guest Relations": "https://chat.whatsapp.com/FmzXrnbwfUaIX8EXHFBxso",
    "Participant Relations": "https://chat.whatsapp.com/LHXFi7uI5f5LT7eSPhep9f",
    "Promotion": "https://chat.whatsapp.com/FfW7M2gRg6tDZJYacEZIdu",
    "Pseudowar": "https://chat.whatsapp.com/LrDmBch5lGzK2P4b3KVosZ",
    "Resources": "https://chat.whatsapp.com/G1hHHItTAeCB699eER16lX",
    "SOP Team": "https://chat.whatsapp.com/FzkPTiaRoqG3XY0eKnGd4u",
    "Security": "https://chat.whatsapp.com/LENQf6zhudz1TGsfmE45OP",
    "Seminar": "https://chat.whatsapp.com/HZ8O61xkQVC8oN6JyEC8zA",
    "Speed Progamming": "https://chat.whatsapp.com/DAB3dHMrHG3IwkQTJd3kq6",
    "UI/UX Competition": "https://chat.whatsapp.com/KGIzHfwQVvCDfL9057sw8r",
    "Web Development Competition": "https://chat.whatsapp.com/KgyWzpCdXyPCvuHIC7Phf0",
    "Web Development": "https://chat.whatsapp.com/IQfpOow7ALbE6uTHeT3KNN",
    "Animations": "https://chat.whatsapp.com/JEqRPIPv2Y6CEMlR3ZtP8s",
    "Automation": "https://chat.whatsapp.com/JtP7jNhsmrHLYX7sSGEmhA",
    "Content": "https://chat.whatsapp.com/L8c34hDffAJ9GrchVZwsqt",
    "Design": "https://chat.whatsapp.com/K48dfGNc0cmEq2UZkdNWkd",
    "Marketing": "https://chat.whatsapp.com/FFoddqocOvZ8rJZV7eVwAA",
    "Media": "https://chat.whatsapp.com/HpXcVqdnK6ALrkEa1XYHa7",
}

print("================================================")
print("         DEV DAY MEMBER MAILS MANAGER")
print("================================================\n")

# TODO replace with master csv in case of cron jobs
dataCsvPath = easygui.fileopenbox()

unsentMailsCsvPath = "unsentRecords.csv"
sentMailsCsvPath = "sentRecords.csv"
processLogFilePath = "processLogs.log"

# Opening Log File
processLogFile = open(processLogFilePath, "a")

unsentRecords = readRecordsFromCsv(dataCsvPath)
sentRecords = readRecordsFromCsv(sentMailsCsvPath)

totalRecords = len(unsentRecords)
unsentLength = totalRecords

processLogFile.write(f"{datetime.now()} : PROCESS STARTED\n")

try:

    i = 0

    while i < unsentLength:
        memberData = unsentRecords[i]

        letter = getLetter(
            memberData["FULL NAME"],
            memberData["SELECT ON-DAY TEAM"],
            memberData["SELECT POSITION"],
        )

        groupLink = whatsApplinks[memberData["SELECT ON-DAY TEAM"]]
        htmlContent = getHtmlContent(groupLink)

        # sort the data according to if the mail was sent or not
        if (
            sendPdfAttachmentMail(memberData["Email Address"], letter, htmlContent)
            == True
        ):
            unsentRecords.remove(memberData)
            sentRecords.append(memberData)
            processLogFile.write(f"{datetime.now()} : SUCCESSFULLY SENT TO {memberData["Email Address"]}\n")
            i -= 1
            unsentLength -= 1
        else:
            processLogFile.write(f"{datetime.now()} : FAILED TO SEND TO {memberData["Email Address"]}\n")
        i += 1


except Exception as ex:
    print("[!] AN ERROR OCCOURED:-")
    processLogFile.write(f"""{datetime.now()} : EXCEPTION OCCOURED\n\t{ex}\n""")
    print(ex)

finally:

    print("[+] Writing data to files before exiting...")

    if writeRecordsToCsv(sentRecords, sentMailsCsvPath):
        print("   [+] Sent records written to file")
    else:
        print("   [+] No sent records to write.")

    if writeRecordsToCsv(unsentRecords, unsentMailsCsvPath):
        print("   [+] Unsent records written to file")
    else:
        print("   [+] No unsent records to write.")

    # Logging process end
    processLogFile.write(f"{datetime.now()} : PROCESS ENDED\n")

    # Closing log file
    processLogFile.close()

print("\n\n======== OPERATION SUMMARY ========")
print(f"\nTotal records to send: {totalRecords}")
print(f"\nSuccessful mails: {totalRecords - unsentLength}")
print(f"--> Saved in: {sentMailsCsvPath}")

print(f"\nUnsuccessful mails: {unsentLength}")
print(f"--> Saved in: {unsentMailsCsvPath}")
