// Configuration used when calling the Microsoft Graph API
const graphConfig = {
    // TODO: Replace with your SharePoint site's Graph API URL
    // To find this:
    // 1. Go to https://graph.microsoft.com/v1.0/sites/YOUR_SHAREPOINT_DOMAIN.sharepoint.com
    // 2. Find your site in the response and copy the full site URL from the Graph response
    // Format: "https://graph.microsoft.com/v1.0/sites/yourdomain.sharepoint.com,site-id,web-id"
    // Example: "https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,12345678-1234-1234-1234-123456789012,12345678-1234-1234-1234-123456789012"
    baseSiteUrl: "YOUR_SHAREPOINT_SITE_GRAPH_URL_HERE",
    
    // TODO: Replace with your SharePoint list's GUID
    // To find this:
    // 1. Go to your SharePoint list
    // 2. Click the gear icon -> List settings
    // 3. Look at the URL - the List parameter contains your list ID
    // Or use Graph Explorer: GET https://graph.microsoft.com/v1.0/sites/YOUR_SITE_URL/lists
    // Format: "12345678-1234-1234-1234-123456789012"
    listId: "YOUR_LIST_ID_HERE"
};

// Get the SharePoint user ID for the provided email address
async function getUserIdByName(name, accessToken) {
    if (!name) return null;
    
    const url = `${graphConfig.baseSiteUrl}/lists/User Information List/items?$filter=fields/EMail eq '${encodeURIComponent(name)}'&$select=id`;

    try {
        const resp = await fetch(url, {
            headers: {
                Authorization: `Bearer ${accessToken}`,
                Accept: "application/json",
                // Required when querying non-indexed SharePoint lists
                'Prefer': 'HonorNonIndexedQueriesWarningMayFailRandomly'
            }
        });
        if (!resp.ok) {
            console.error("Graph API error while getting user by name:", await resp.text());
            return null;
        }
        const data = await resp.json();
        if (data.value && data.value.length > 0) {
            return parseInt(data.value[0].id, 10);
        }
        return null;
    } catch (err) {
        console.error(`User lookup failed for "${name}":`, err);
        return null;
    }
}

// Convert a user ID back to the person's display name
async function getUserById(userId, accessToken) {
    if (!userId) return null;
    const url = `${graphConfig.baseSiteUrl}/lists/User Information List/items/${userId}?$select=fields`;

    try {
        const resp = await fetch(url, {
            headers: { Authorization: `Bearer ${accessToken}`, Accept: "application/json" }
        });
        if (!resp.ok) return null;
        const data = await resp.json();
        return data.fields.Title;
    } catch (err) {
        console.error(`User lookup failed (${userId}):`, err);
        return null;
    }
}

// Read all items from the SharePoint list
async function getListItems() {
    const accessToken = await getAccessToken();
    const url = `${graphConfig.baseSiteUrl}/lists/${graphConfig.listId}/items?$expand=fields`;

    const resp = await fetch(url, {
        headers: { Authorization: `Bearer ${accessToken}`, Accept: "application/json" }
    });
    if (!resp.ok) {
        console.error("Graph API error:", await resp.text());
        throw new Error(`Graph error ${resp.status}`);
    }

    const items = (await resp.json()).value;

    // Replace the numeric user ID with a readable name
    for (const item of items) {
        if (item.fields.AssignmentLookupId) {
            const name = await getUserById(item.fields.AssignmentLookupId, accessToken);
            item.fields.Assignment = { Title: name || "" };
        }
    }

    return items;
}