// When the page loads, set up the UI based on whether the user is signed in
document.addEventListener("DOMContentLoaded", updateUI);

// Fetch the current user's first name using Microsoft Graph
async function getUserFirstName() {
    try {
        const accessToken = await getAccessToken();
        const response = await fetch('https://graph.microsoft.com/v1.0/me', {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Accept': 'application/json'
            }
        });
        
        if (response.ok) {
            const userData = await response.json();
            // Prefer the given name but fall back to the display name or email
            return userData.givenName || userData.displayName || sessionStorage.getItem("msalAccount");
        }
    } catch (error) {
        console.error("Error fetching user data:", error);
    }
    
    // Fallback to the email address if the Graph call fails
    return sessionStorage.getItem("msalAccount");
}

// Show or hide sections of the page depending on sign-in status
async function updateUI() {
    const account = sessionStorage.getItem("msalAccount");
    const signInBtn = document.getElementById("sign-in");
    const signOutBtn = document.getElementById("sign-out");
    const userInfo  = document.getElementById("user-info");
    const content   = document.getElementById("content");

    if (account) {
        signInBtn.style.display  = "none";
        signOutBtn.style.display = "block";
        userInfo.style.display   = "block";
        content.style.display    = "block";
        
        // Get and display the user's first name
        const firstName = await getUserFirstName();
        document.getElementById("user-name").innerText = firstName;
        
        displayItems();
    } else {
        signInBtn.style.display  = "block";
        signOutBtn.style.display = "none";
        userInfo.style.display   = "none";
        content.style.display    = "none";
    }
}

// Render the SharePoint list items as an HTML table
async function displayItems() {
    const items = await getListItems();
    const container = document.getElementById("item-list");

    let html = `<table><tr>
        <th>Title</th>
        <th>Description</th>
        <th>Assignment</th>
        <th>Status</th>
        <th>Size</th>
        <th>Link</th>
        <th>Complete Year</th>
        <th>Location</th>
    </tr>`;

    // Build a table row for each item
    items.forEach(item => {
        const f = item.fields;
        const assignment = f.Assignment ? f.Assignment.Title : "";
        const size       = f.Cost ? f.Cost.toLocaleString() : "";
        const linkHtml   = f.Link ? `<a href="${f.Link.Url}" target="_blank">Link</a>` : "";
        const year       = f.Complete ? new Date(f.Complete).getFullYear() : "";

        html += `<tr>
            <td>${f.Title || ""}</td>
            <td>${f.Description || ""}</td>
            <td>${assignment}</td>
            <td>${f.Status || ""}</td>
            <td>${size}</td>
            <td>${linkHtml}</td>
            <td>${year}</td>
            <td>${f.Location || ""}</td>
        </tr>`;
    });

    html += "</table>";
    container.innerHTML = html;
}
