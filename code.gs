function onOpen() {
  var ui = DocumentApp.getUi();
  ui.createMenu('Requirements')
    .addItem('Update Requirement Links', 'updateRequirementLinks')
    .addItem('Generate Summary Table', 'insertSummaryTable')
    .addSeparator()
    .addItem('Show Metadata Sidebar', 'showSidebar')  // Menu item to show the sidebar
    .addToUi();
}

function updateRequirementLinks() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const requirementKeysPattern = /REQ-\d{3}/g;
  let match;

  while ((match = requirementKeysPattern.exec(body.getText())) !== null) {
    const reqKey = match[0];
    const url = getRequirementUrl(reqKey);
    const rangeElement = body.findText(reqKey);

    if (rangeElement !== null) {
      const startOffset = rangeElement.getStartOffset();
      const endOffset = rangeElement.getEndOffsetInclusive();
      const textElement = rangeElement.getElement().asText();
      textElement.setLinkUrl(startOffset, endOffset, url);
    }
  }
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Requirement Metadata')
    .setWidth(300);
  DocumentApp.getUi().showSidebar(html);
}

function getRequirementUrl(reqKey) {
  var apiUrl = 'https://ww1.requirementyogi.cloud/nuitdelinfo/search';
  var params = {
    'method': 'get',
    'muteHttpExceptions': true
  };
  var response = UrlFetchApp.fetch(apiUrl + '?key=' + encodeURIComponent(reqKey), params);
  var statusCode = response.getResponseCode();
  var url = '';
  if (statusCode === 200) {
    var jsonResponse = JSON.parse(response.getContentText());
    if (jsonResponse.results && jsonResponse.results.length > 0) {
      url = jsonResponse.results[0].canonicalUrl;
    }
  } else {
    console.error('Error fetching URL for ' + reqKey + ': ' + response.getContentText());
  }
  return url;
}

function loadMetadata() {
  var reqKeyInput = document.getElementById('reqKeyInput');
  if (reqKeyInput && reqKeyInput.value) {
    var reqKey = reqKeyInput.value.trim();
    google.script.run.withSuccessHandler(displayMetadata)
                     .withFailureHandler(handleError)
                     .getRequirementMetadata(reqKey); // Pass reqKey as an argument
  } else {
    alert('Please enter a Requirement Key.');
  }
}

  function displayMetadata(metadata) {
    var container = document.getElementById('metadata-container');
    if (metadata && metadata.title) {
      container.innerHTML = '<strong>Title:</strong> ' + metadata.title +
                            '<br><strong>Description:</strong> ' + metadata.description;
    } else {
      container.innerHTML = 'No metadata found for the given REQ key.';
    }
  }

  function handleError(error) {
    var container = document.getElementById('metadata-container');
    container.innerHTML = 'Error: ' + error.message;
  }

function getRequirementMetadata(reqKey) {
  var apiUrl = 'https://ww1.requirementyogi.cloud/nuitdelinfo/metadata';
  var params = {
    'method': 'get',
    'muteHttpExceptions': true
  };
  try {
    var response = UrlFetchApp.fetch(apiUrl + '?key=' + encodeURIComponent(reqKey), params);
    var metadata = {};
    if (response.getResponseCode() === 200) {
      var jsonResponse = JSON.parse(response.getContentText());
      Logger.log(jsonResponse);  // Log the entire response for debugging
      if (jsonResponse.results && jsonResponse.results.length > 0) {
        metadata = jsonResponse.results[0]; // Assuming metadata is in results[0]
      }
    } else {
      console.error('Error fetching metadata for ' + reqKey + ': ' + response.getContentText());
    }
    return metadata;
  } catch (e) {
    Logger.log('Error fetching metadata: ' + e.toString());
    return {}; // Return an empty object if there was an error
  }
}


function insertSummaryTable() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var requirementKeys = body.getText().match(/REQ-\d{3}/g);
  if (requirementKeys) {
    requirementKeys = [...new Set(requirementKeys)];
    var tableTitle = body.appendParagraph('Requirement Keys Summary');
    tableTitle.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    var table = body.appendTable();
    var headerRow = table.appendTableRow();
    headerRow.appendTableCell('Requirement Key').setBackgroundColor('#dddddd');
    headerRow.appendTableCell('URL').setBackgroundColor('#dddddd');
    requirementKeys.forEach(function(key) {
      var url = getRequirementUrl(key);
      var row = table.appendTableRow();
      row.appendTableCell(key);
      var linkCell = row.appendTableCell('');
      var textRange = linkCell.editAsText();
      textRange.setText(url);
      textRange.setLinkUrl(url);
    });
  } else {
    DocumentApp.getUi().alert('No requirement keys found.');
  }
}

// No need to manually run the setUp function since onOpen is a trigger function
// that runs automatically when the document is opened.
