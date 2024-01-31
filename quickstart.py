import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1a6N6NadfgMou-lWNDkX4-5yuBSkZzOIqISToJft6rXw"
SAMPLE_RANGE_NAME = "Página1!C4:F27"


def main():
    creds = None

    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port = 0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("sheets", "v4", credentials = creds)

        # Read Google Sheets Cels
        sheet = service.spreadsheets()
        result = (sheet.values().get(
            spreadsheetId = SAMPLE_SPREADSHEET_ID,
            range = SAMPLE_RANGE_NAME
        ).execute())

        # Get Absences and Means
        rows = result['values']
        absences = list()
        means = list()
        sum = loop = 0

        for value in rows:
            for num in value:
                if loop == 0:
                    absences.append(int(num))
                else:
                    sum += int(num)
                loop += 1
                if loop == 4:
                    means.append(round(sum / 3))
                    loop = sum = 0

        # Situations Logic
        student = list(zip(absences, means))
        situations = list()

        for tuple in student:
            abs_percent = round((tuple[0] * 100) / 60)
            if abs_percent > 25:
                situations.append(['Reprovado por Falta', 0])
            else:
                if tuple[1] < 50:
                    situations.append(['Reprovado por Nota', 0])
                elif tuple[1] < 70:
                    naf = tuple[1] - 10
                    situations.append(['Exame Final', naf])
                else:
                    situations.append(['Aprovado', 0])

        # Write Google Sheet Cels
        (sheet.values().update(
            spreadsheetId = SAMPLE_SPREADSHEET_ID, range = 'G4',
            valueInputOption = 'USER_ENTERED', body = {'values': situations}
        ).execute())

    except HttpError as err:
        print(err)


if __name__ == "__main__":
    main()
