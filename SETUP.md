# ATL Integration Hub - GitHub Pages Setup

## Architecture

```
User → SharePoint Page (iframe) → GitHub Pages → Power Automate Flow → SharePoint List
```

## Step 1: Create the Power Automate Flow

1. Go to https://make.powerautomate.com
2. Click **My flows** → **Import** → **Import Package (Legacy)**
3. Upload `flows/ATL-DataProxy-Flow.json`
4. Configure the SharePoint connection
5. Save and turn on the flow
6. Copy the HTTP POST URL from the trigger

**Alternative - Manual Creation:**

1. Create a new **Instant cloud flow**
2. Trigger: **When a HTTP request is received**
3. Method: POST
4. Request Body JSON Schema:
```json
{
    "type": "object",
    "properties": {
        "listName": { "type": "string" },
        "select": { "type": "string" },
        "filter": { "type": "string" },
        "top": { "type": "integer" },
        "orderby": { "type": "string" }
    },
    "required": ["listName"]
}
```
5. Add **SharePoint - Get items** action:
   - Site: `https://chamberlaingroup.sharepoint.com/sites/PrincipalGTMStrategy-InternalUseOnly-ATLIntegrationProject`
   - List: `@{triggerBody()?['listName']}`
   - Top: `@{if(equals(triggerBody()?['top'], null), 500, triggerBody()?['top'])}`
   - Filter: `@{triggerBody()?['filter']}`
   - Order By: `@{triggerBody()?['orderby']}`

6. Add **Response** action:
   - Status: 200
   - Headers: `Access-Control-Allow-Origin: *`
   - Body:
```json
{
    "items": @{body('Get_items')?['value']},
    "count": @{length(body('Get_items')?['value'])},
    "timestamp": "@{utcNow()}"
}
```

7. Save and copy the HTTP POST URL

## Step 2: Configure spApi.js

Edit `spApi.js` and set the Flow URL:

```javascript
// Line 27 - set your Flow URL
flowProxyUrl: 'https://prod-XX.westus.logic.azure.com:443/workflows/YOUR-FLOW-ID/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=YOUR-SIG',
```

Or configure at runtime:
```javascript
spApi.setFlowProxyUrl('YOUR-FLOW-URL-HERE');
```

## Step 3: Deploy to GitHub Pages

```bash
cd /mnt/d/Dev/atl/output/github-pages
git add .
git commit -m "Initial ATL Dashboard deployment"
gh repo create atl-dashboards --public --source=. --push
```

Enable GitHub Pages:
1. Go to repo Settings → Pages
2. Source: Deploy from branch
3. Branch: main, folder: / (root)
4. Save

Your dashboards will be live at: `https://YOUR-USERNAME.github.io/atl-dashboards/`

## Step 4: Create SharePoint Embed Pages

For each dashboard, create a SharePoint page with an Embed web part:

1. Go to SharePoint site → Site Pages → New → Page
2. Add **Embed** web part
3. Paste iframe code:

```html
<iframe
    src="https://YOUR-USERNAME.github.io/atl-dashboards/index.html"
    width="100%"
    height="800"
    frameborder="0"
    style="border: none;">
</iframe>
```

4. Publish the page

## Page URLs

| Dashboard | GitHub Pages URL |
|-----------|------------------|
| Landing | `/index.html` |
| Status Hub | `/ATL_Status_Hub.html` |
| Gantt Chart | `/ATL_Project_Gantt.html` |
| Task Detail | `/ATL_Task_Detail.html` |
| My Tasks | `/ATL_My_Tasks.html` |
| Team Workload | `/ATL_Team_Workload.html` |
| Budget Tracker | `/ATL_Budget_Tracker.html` |
| Blockers | `/ATL_Blocker_Dashboard.html` |
| Reports | `/ATL_Reports.html` |
| Admin | `/ATL_Admin.html` |
| Milestones | `/ATL_Milestone_Timeline.html` |

## Testing

1. Open GitHub Pages URL directly
2. Open browser DevTools (F12) → Console
3. Look for `[spApi]` messages
4. Should see: `Flow proxy returned 172 items from ATL_Project_Plan.v21`

## Troubleshooting

| Issue | Solution |
|-------|----------|
| CORS error | Verify Flow has `Access-Control-Allow-Origin: *` header |
| 401 Unauthorized | Flow connection needs reauthorization |
| 404 List not found | Check list name spelling in request |
| Empty data | Check Flow run history for errors |

## Security Notes

- The Flow HTTP trigger URL contains a SAS signature - treat as secret
- Flow runs under your credentials - users see data you can see
- No write operations - read-only data access
