Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) {
        if (info.platform === Office.PlatformType.PC || info.platform === Office.PlatformType.Web) {
            // Office is ready
            console.log("Outlook Add-in is ready on PC or Web");
            try {
                loadEmails(); // Call to load emails
            } catch (error) {
                console.error("Error calling loadEmails:", error);
            }
        } else {
            console.warn("Outlook Add-in is running on an unsupported platform");
        }
    }
});

function loadEmails() {
    console.log("Loading emails...");
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log("Access token received");
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
                console.log("Emails fetched successfully");
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
    console.log("Allowing drop on target:", event.target);
}

function drag(event) {
    console.log("Dragging item:", event.target.id);
    event.dataTransfer.setData("text", event.target.id);
}

function drop(event) {
    event.preventDefault();
    console.log("Dropping item on target:", event.target);
    const data = event.dataTransfer.getData("text");
    const emailElement = document.getElementById(data);
    if (event.target.id === 'dropZone') {
        event.target.appendChild(emailElement);
    } else {
        console.warn("Drop attempted on incorrect target");
    }
}
