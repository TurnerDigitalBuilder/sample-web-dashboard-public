// Handle submission of the "Add item" form. This creates a new item in the
// SharePoint list defined in `graph.js`.
async function addItem(event) {
    // Prevent the browser from reloading the page when the form is submitted
    event.preventDefault();

    // Grab all form fields so we can build the item object
    const title = document.getElementById('title').value;
    const description = document.getElementById('description').value;
    const assignment = document.getElementById('assignment').value;
    const status = document.getElementById('status').value;
    const size = document.getElementById('size').value;
    const completeYear = document.getElementById('complete_year').value;
    const location = document.getElementById('location').value;

    // We need a Graph access token in order to talk to SharePoint
    const accessToken = await getAccessToken();

    // Convert the person's email address into a numeric user ID that
    // SharePoint understands. If the user cannot be found we stop here.
    let assignmentId = null;
    if (assignment) {
        assignmentId = await getUserIdByName(assignment, accessToken);
        if (!assignmentId) {
            console.error("Could not find user:", assignment);
            alert("Error: Assigned user not found.");
            return;
        }
    }

    // Build up the payload that will be sent to Microsoft Graph. The
    // property names must match the columns in the SharePoint list.
    const newItemFields = {
        Title: title,
        Description: description,
        Status: status,
        Cost: parseFloat(size) || 0,
        Complete: new Date(completeYear, 0, 1).toISOString(),
        Location: location
    };

    // Only include the assignment field if a user was selected
    if (assignmentId) {
        newItemFields.AssignmentLookupId = assignmentId;
    }

    try {
        // Use the helper in graph.js to create the item via Graph
        await addListItem({ fields: newItemFields });

        // Refresh the list in the UI so the new item appears
        displayItems();

        // Clear the form inputs ready for the next entry
        event.target.reset();
    } catch (error) {
        console.error("Error adding item:", error);
    }
}
