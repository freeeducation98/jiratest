import requests
from requests.auth import HTTPBasicAuth
import pandas as pd

# ========== CONFIGURE THIS ==========
JIRA_URL = "JIRA_URL"
USERNAME = "your-username"
PASSWORD = "your-password"
JQL_QUERY = "project = YOURPROJECT ORDER BY created DESC"
FIELDS = "summary,status,assignee,created,updated"
EXCEL_OUTPUT = "jira_issues_export.xlsx"
# ====================================

MAX_RESULTS = 100
start_at = 0
all_issues = []

print("Starting export from Jira...")

while True:
    api_url = f"{JIRA_URL}/rest/api/2/search"
    params = {
        'jql': JQL_QUERY,
        'startAt': start_at,
        'maxResults': MAX_RESULTS,
        'fields': FIELDS
    }

    response = requests.get(api_url, params=params, auth=HTTPBasicAuth(USERNAME, PASSWORD))

    if response.status_code != 200:
        print(f"Failed to fetch data: {response.status_code}")
        print(response.text)
        break

    data = response.json()
    issues = data.get("issues", [])

    if not issues:
        print("No more issues to fetch.")
        break

    for issue in issues:
        fields = issue["fields"]
        all_issues.append({
            "Key": issue["key"],
            "Summary": fields.get("summary"),
            "Status": fields.get("status", {}).get("name"),
            "Assignee": fields.get("assignee", {}).get("displayName") if fields.get("assignee") else None,
            "Created": fields.get("created"),
            "Updated": fields.get("updated")
        })

    print(f"Fetched {len(issues)} issues (StartAt: {start_at})")
    start_at += MAX_RESULTS

# Save to Excel
df = pd.DataFrame(all_issues)
df.to_excel(EXCEL_OUTPUT, index=False)
print(f"\n✅ Export complete! {len(df)} issues written to {EXCEL_OUTPUT}")
