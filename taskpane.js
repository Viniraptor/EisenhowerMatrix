Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) {
        // Code to execute when the task pane is ready
        loadEmails();
    }
});

function loadEmails() {
    // Load unread emails into a draggable list
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const accessToken = result.value;
            const mailboxUrl = Office.context.mailbox.restUrl;
            const requestUrl = `${mailboxUrl}/v2.0/me/mailfolders/inbox/messages?$filter=isRead eq false`;

            fetch(requestUrl, {
                headers: {
                    'Authorization': `Bearer ${accessToken}`
                }
            })
            .then(response => response.json())
            .then(data => {
                const emailList = data.value;
                emailList.forEach((email) => {
                    const emailElement = document.createElement('div');
                    emailElement.className = "email-item";
                    emailElement.draggable = true;
                    emailElement.ondragstart = drag;
                    emailElement.id = email.id;
                    emailElement.innerText = `${email.subject}`;
                    document.getElementById('emailsContainer').appendChild(emailElement);
                });
            })
            .catch(error => {
                console.error('Error fetching unread emails:', error);
            });
        } else {
            console.error('Error getting callback token:', result.error);
        }
    });
}

function allowDrop(event) {
    event.preventDefault();
    event.target.style.border = "2px solid #00f"; // Highlight the drop zone
}

function drag(event) {
    event.dataTransfer.setData("text", event.target.id);
}

function drop(event) {
    event.preventDefault();
    event.target.style.border = "2px dashed #ccc"; // Reset the border style
    const data = event.dataTransfer.getData("text");
    const emailElement = document.getElementById(data);
    document.getElementById('dropZone').appendChild(emailElement);
}

// Add this to your HTML file or task pane:
// <div id="emailsContainer" style="padding: 10px; border: 1px solid #ddd;">Drag emails here to view</div>
// <div id="dropZone" ondrop="drop(event)" ondragover="allowDrop(event)" style="border: 2px dashed #ccc; min-height: 200px; margin-top: 20px;">Drop emails here</div>
