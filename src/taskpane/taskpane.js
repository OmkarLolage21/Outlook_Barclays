/*
* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
* See LICENSE in the project root for license information.
*/
/* global document, Office */


// Ik, this coding practice is kinda disturbing, but dealt with too many issues on CORS,
// accepting access token for a certain user, etc. given the time constraint had to do this/
// Focus was to get the basic app working
// global variables that can be accessed by all functions
var config = {
  accessToken: "eyJ0eXAiOiJKV1QiLCJub25jZSI6IlZ0dGIzaFpaUHFhZ1lHR016U0dGNVVQdVpkQ1JNNWptUm91d2pYRG5lNkEiLCJhbGciOiJSUzI1NiIsIng1dCI6InEtMjNmYWxldlpoaEQzaG05Q1Fia1A1TVF5VSIsImtpZCI6InEtMjNmYWxldlpoaEQzaG05Q1Fia1A1TVF5VSJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC8zMTVlMzJhNC1jMTAwLTQ1YTktODUxNS01YWU1M2Y4YjE2MTkvIiwiaWF0IjoxNzEyMzg0MzI3LCJuYmYiOjE3MTIzODQzMjcsImV4cCI6MTcxMjQ3MTAyNywiYWNjdCI6MCwiYWNyIjoiMSIsImFpbyI6IkFUUUF5LzhXQUFBQXk4dEw1bXZDbFRicWMzME5mZTFra1JYQ3U1VjJBcWVDOVVPYjVyRk5FbFRHblFpMjFralVLT1dqSjdkRDRadEEiLCJhbXIiOlsicHdkIl0sImFwcF9kaXNwbGF5bmFtZSI6IkdyYXBoIEV4cGxvcmVyIiwiYXBwaWQiOiJkZThiYzhiNS1kOWY5LTQ4YjEtYThhZC1iNzQ4ZGE3MjUwNjQiLCJhcHBpZGFjciI6IjAiLCJmYW1pbHlfbmFtZSI6ImxvbGFnZSIsImdpdmVuX25hbWUiOiJvbWthciIsImlkdHlwIjoidXNlciIsImlwYWRkciI6IjE0OC42NC40LjgxIiwibmFtZSI6Im9ta2FyIGxvbGFnZSIsIm9pZCI6ImMwMmQyYTg5LWE3NmYtNDBlNS05NDU5LTBmZjkxNGMyOWUyZiIsInBsYXRmIjoiMyIsInB1aWQiOiIxMDAzMjAwMzZBRTI3NkY1IiwicmgiOiIwLkFTc0FwREplTVFEQnFVV0ZGVnJsUDRzV0dRTUFBQUFBQUFBQXdBQUFBQUFBQUFEQ0FMcy4iLCJzY3AiOiJBUElDb25uZWN0b3JzLlJlYWQuQWxsIEFQSUNvbm5lY3RvcnMuUmVhZFdyaXRlLkFsbCBNYWlsLlJlYWQgTWFpbC5SZWFkQmFzaWMgTWFpbC5SZWFkV3JpdGUgTWFpbC5TZW5kIE1haWxib3hTZXR0aW5ncy5SZWFkIG9wZW5pZCBwcm9maWxlIFVzZXIuUmVhZCBlbWFpbCIsInN1YiI6Im9GOFlTa2NUZlc4WGRmY192UDYxLVFzeGJnWHZQNmtXRU51dGY3VmZpejgiLCJ0ZW5hbnRfcmVnaW9uX3Njb3BlIjoiQVMiLCJ0aWQiOiIzMTVlMzJhNC1jMTAwLTQ1YTktODUxNS01YWU1M2Y4YjE2MTkiLCJ1bmlxdWVfbmFtZSI6Im9ta2FyTG9sYWdlQGhhY2tvaGlyZS5vbm1pY3Jvc29mdC5jb20iLCJ1cG4iOiJvbWthckxvbGFnZUBoYWNrb2hpcmUub25taWNyb3NvZnQuY29tIiwidXRpIjoieXlsam16UUdvVWlOV29TUThCZ05BQSIsInZlciI6IjEuMCIsIndpZHMiOlsiNjJlOTAzOTQtNjlmNS00MjM3LTkxOTAtMDEyMTc3MTQ1ZTEwIiwiYjc5ZmJmNGQtM2VmOS00Njg5LTgxNDMtNzZiMTk0ZTg1NTA5Il0sInhtc19jYyI6WyJDUDEiXSwieG1zX3NzbSI6IjEiLCJ4bXNfc3QiOnsic3ViIjoicmtlRFN4eUZZamRLcG8taHJRampQd1VpNmJzVzJJUVRWMjBmSDA5WWo4ZyJ9LCJ4bXNfdGNkdCI6MTcxMTcwMjUxN30.bxzOCeluegS7RziGnn_T78OTOU_gViLq_yLoLIjAUVN_aiHPJrHbXPS2f1EVP_LMZNx_PeP8N_bei5eLv8AkzHBpMzUnKFDpS9aA8RQ4EqxmswsxnJng_87Ekz50Ab41-zP_sAh03REl-7JLD1q0LhRhX3uIfPZmreZANaLcS8cIE5tHn2hDhG7d15uA17_sjbNPcFApTRdP1ER5f_ikA1hL5TU4uufxCvqd7OhdIU39KZ2hlcnfXeogRRPPdc_EpZ6Z8TsxIkXsUhVjyPNhioiCGGzWZnXpuh2qWQqRFGtQDoUOHxOlj5VwjbVThcp8w_QP3U1kS5YX1wN42E0mog"
};

const emails = [];

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // document.getElementById("sideload-msg").style.display = "none";
    // document.getElementById("app-body").style.display = "flex";
    // document.getElementById("email-container").onclick = showMessage;

    //  document.getElementById('getAccessToken-button').addEventListener('click', getAccessToken);
    document.getElementById('getMails-button').addEventListener('click', getMessages);
    document.getElementById('forwardEmails-button').addEventListener('click', forwardEmails);
  }
});

//----------------------------------------------------------------------------
// Function to log messages
function logMessage(message) {
  var emailContainer = document.getElementById('email-container');
  emailContainer.innerText = message;
}

// Function to hide message dialog
function hideMessage() {
  var messageDialog = document.getElementById('messageDialog');
  messageDialog.style.display = 'none';
}

async function getMessages() {
  // Clear the emails array to overwrite it with new emails
  emails.length = 0; // This line clears the array
 
  // Define the endpoint to fetch emails
  const baseEndpoint = 'https://graph.microsoft.com/v1.0/me/messages?$select=id,subject,sender';
 
  // Get the selected value from the dropdown
  const selectedCategory = document.getElementById('selectEmail').value;
 
  // Construct the endpoint dynamically based on the selected category
  const endpoint = selectedCategory ? `${baseEndpoint}&$filter=categories/any(c:c eq '${selectedCategory}')` : baseEndpoint;
 
  // Define the access token
  const accessToken = config.accessToken;
 
  // Define headers with the access token
  const headers = {
     'Authorization': `Bearer ${accessToken}`
  };
 
  try {
     // Make a GET request to fetch emails
     const response = await fetch(endpoint, {
       method: 'GET',
       headers: headers
     });
 
     // Check if the response is successful
     if (!response.ok) {
       throw new Error(`Failed to fetch emails: ${response.statusText}`);
     }
 
     // Extract the response body as JSON
     const responseBody = await response.json();
 
     // Iterate over each email object and extract required information
     responseBody.value.forEach(email => {
       const emailObj = {
        id: email.id, 
         subject: email.subject,
         sender: email.sender.emailAddress.name
       };
 
       // Push the email object to the emails array
       emails.push(emailObj);
     });
 
     // Render the emails in the task pane
     renderEmails(emails);
  } catch (error) {
     console.error('Error fetching emails:', error.message);
  }
 }
 

function renderEmails(emails) {
  // Get the task pane container element
  var taskPaneContainer = document.getElementById('task-pane-container');

  // Clear previous content in the task pane container
  taskPaneContainer.innerHTML = '';

  // Iterate over each email object and render a custom flexbox for each email
  emails.forEach(email => {
    // Create a new div element to hold the email details with flexbox layout
    var emailDiv = document.createElement('div');
    emailDiv.className = 'email';
    emailDiv.style.display = 'flex';
    emailDiv.style.alignItems = 'center';
    emailDiv.style.marginBottom = '5px'; // Reduced marginBottom for spacing
    emailDiv.style.height = '50px'; // Set the height to 50px (adjust as needed)

    // Create a div for the checkbox (left part, occupying 10% width)
    var checkboxDiv = document.createElement('div');
    checkboxDiv.style.width = '10%';
    checkboxDiv.style.height = '50px%'; // Increased height for the checkbox
    checkboxDiv.style.display = 'flex';
    checkboxDiv.style.justifyContent = 'center'; // Center the checkbox vertically

    // Create a checkbox element for the email
    var checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.id = email.id; // Use the email ID as the checkbox ID

    // Append the checkbox to the checkbox div
    checkboxDiv.appendChild(checkbox);

    // Create a div for the email details (right part)
    var detailsDiv = document.createElement('div');
    detailsDiv.style.width = '90%';
    detailsDiv.style.paddingLeft = '10px'; // Add some left padding for spacing

    // Create a div for the sender
    var senderDiv = document.createElement('div');
    senderDiv.innerHTML = '<strong>' + email.sender + '</strong>'; // Bold sender name
    senderDiv.style.marginBottom = '2px'; // Reduced marginBottom for spacing
    senderDiv.style.overflow = 'hidden'; // Hide overflow content
    senderDiv.style.textOverflow = 'ellipsis'; // Show ... for overflow content
    senderDiv.style.whiteSpace = 'nowrap'; // Display sender name in single line

    // Calculate the available width for sender and subject based on task pane width
    var availableWidth = taskPaneContainer.offsetWidth * 0.9; // 90% of task pane width
    var subjectWidth = availableWidth * 0.3; // 30% of available width for subject

    // Create a temporary div to measure the width of the sender's content
    var tempDiv = document.createElement('div');
    tempDiv.innerHTML = senderDiv.innerHTML;
    tempDiv.style.position = 'absolute';
    tempDiv.style.visibility = 'hidden';
    document.body.appendChild(tempDiv);
    var senderWidth = tempDiv.offsetWidth; // Get the width of the sender's content

    // Truncate sender name if it exceeds the available width
    if (senderWidth > availableWidth * 0.7) {
      senderDiv.innerHTML = '<strong>' + truncateText(email.sender, (availableWidth * 0.7) - 10) + '</strong>'; // Adjusting for padding
    }

    // Truncate sender name if it exceeds the available width
    senderDiv.innerHTML = '<strong>' + truncateText(email.sender, senderWidth) + '</strong>';

    // Create a div for the subject
    var subjectDiv = document.createElement('div');
    subjectDiv.innerHTML = email.subject; // Display subject without truncation
    subjectDiv.style.overflow = 'hidden'; // Hide overflow content
    subjectDiv.style.textOverflow = 'ellipsis'; // Show ... for overflow content

    // Append sender and subject divs to the details div
    detailsDiv.appendChild(senderDiv);
    detailsDiv.appendChild(subjectDiv);

    // Append checkbox div and details div to the email div
    emailDiv.appendChild(checkboxDiv);
    emailDiv.appendChild(detailsDiv);

    // Append the email div to the task pane container
    taskPaneContainer.appendChild(emailDiv);
  });
}


// Function to truncate text if longer than a specified width
function truncateText(text, maxWidth) {
  var truncatedText = text;
  var ellipsis = '...';
  if (text.length > maxWidth) {
    truncatedText = text.slice(0, maxWidth - ellipsis.length) + ellipsis;
  }
  return truncatedText;
}

async function forwardEmails() {
  // Get the checkboxes for the selected emails
  const checkboxes = document.querySelectorAll('input[type="checkbox"]:checked');

  // Check if any email is selected
  if (checkboxes.length === 0) {
    console.log('Please select at least one email to forward.');
    return;
  }

  // Get the selected recipient email address and name
  const recipientEmail = document.getElementById('recipients-dropdown').value;
  const recipientName = document.getElementById('recipients-dropdown').options[document.getElementById('recipients-dropdown').selectedIndex].text;
  console.log('Selected recipient email:', recipientEmail); // Log the selected recipient email
  console.log('Selected recipient name:', recipientName); // Log the selected recipient name

  // Check if a recipient is selected
  if (!recipientEmail) {
    console.log('Please select a recipient email address.');
    return;
  }

  // Define headers with the access token
  const headers = {
    'Authorization': `Bearer ${config.accessToken}`,
    'Content-Type': 'application/json'
  };
  console.log('Headers:', headers); // Log the headers

  try {
    // Iterate over each selected email and forward it
    await Promise.all(Array.from(checkboxes).map(async checkbox => {
      // Find the corresponding email object in the emails array
      const emailObj = emails.find(email => email.id === checkbox.id);
      if (!emailObj) {
        throw new Error(`Email not found for ID: ${checkbox.id}`);
      }

      const emailId = emailObj.id;
      const forwardEndpoint = `https://graph.microsoft.com/v1.0/me/messages/${emailId}/microsoft.graph.forward`;
      console.log('Forward endpoint:', forwardEndpoint); // Log the forward endpoint

      // Define the request body for forwarding
      const requestBody = JSON.stringify({
        "comment": "FYI",
        "toRecipients": [
          {
            "emailAddress": {
              "address": recipientEmail,
              "name": recipientName // Use the recipient's name from the dropdown
            }
          }
        ]
      });
      console.log('Request body:', requestBody); // Log the request body

      // Make a POST request to forward the email
      const forwardResponse = await fetch(forwardEndpoint, {
        method: 'POST',
        headers: headers,
        body: requestBody
      });

      // Check if the forwarding was successful
      if (!forwardResponse.ok) {
        const errorResponse = await forwardResponse.json(); // Attempt to parse the error response
        console.log('Error response:', errorResponse); // Log the error response
        throw new Error(`Failed to forward email: ${forwardResponse.statusText}. Error details: ${JSON.stringify(errorResponse)}`);
      }
    }));

    console.log('Emails forwarded successfully!');
  } catch (error) {
    console.error('Error forwarding emails:', error.message);
  }
}
