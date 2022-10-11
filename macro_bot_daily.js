let webhookURL = "";
let webappURL = "";

function loadVariables()
{
  let spreadsheet = SpreadsheetApp.getActive();
  let sheet = spreadsheet.getActiveSheet();
  let range = sheet.getRange("F2:G2").activate();
  let values = range.getValues();

  webhookURL = values[0][0];
  webappURL = values[0][1];
}

function doGet(e) 
{
  setFacilitator();
}

function doPost(e) 
{
  setFacilitator();
}

function setFacilitator()
{
  loadVariables();

  let facilitator = getFacilitator();

  let spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D2').activate();
  spreadsheet.getCurrentCell().setValue(facilitator.name);
  spreadsheet.getRange('D3').activate();

  setNewWeight(facilitator);

  let payload = buildPayload(facilitator.name);
  
  sendAlert(payload);

  updateLastExecution();
}

function setNewWeight(facilitator)
{
  let data = getPeopleData();

  for (let row in data)
  {
    if (data[row][0] == facilitator.id)
    {
      let newWeight = data[row][2] + 1;

      let rowNumber = facilitator.id + 1;

      let spreadsheet = SpreadsheetApp.getActive();
      spreadsheet.getRange("C" + rowNumber).activate();
      spreadsheet.getCurrentCell().setValue(newWeight);
      spreadsheet.getRange('A1').activate();
    }
  }
}

function getFacilitator() 
{
  let arPeople = orderPeopleArrayByWeight(createPeopleArray());
  let facilitator;

  if (validatePeopleArray(arPeople))
  {
    facilitator = arPeople[0];

    let lighterWeight = arPeople[0].weight;
    let arPeopleLighterWight = [];

    for (let x = 0; x < arPeople.length; x++)
    {
      if (arPeople[x].weight == lighterWeight)
        arPeopleLighterWight.push(arPeople[x]);
    }

    if (validatePeopleArray(arPeopleLighterWight))
      facilitator = orderPeopleArrayRandomly(arPeopleLighterWight)[0];

    return facilitator;
  }
};

function validatePeopleArray(arr)
{
  if (arr == null || arr.length == 0)
  {
    Logger.log("Array de pessoas vazio");
    return false;
  }

  return true;
}

function createPeopleArray() 
{
  let data = getPeopleData();
  let arPeople = [];

  for (let row in data) 
  {
    if (data[row][0] != "") 
    {
        let ojbPerson = 
        {
          id: data[row][0],
          name: data[row][1],
          weight: data[row][2]
        };

      arPeople.push(ojbPerson);  
    }
  };

  return arPeople;
}

function orderPeopleArrayByWeight(arPeople) 
{
  return arPeople.sort(compareWeight);
}

function orderPeopleArrayRandomly(arPeople) 
{ 
  let currentIndex = arPeople.length;

  while (currentIndex != 0) 
  {
    let randomIndex = Math.floor(Math.random() * currentIndex);

    currentIndex--;

    let personAux = arPeople[currentIndex];

    arPeople[currentIndex] = arPeople[randomIndex];
    arPeople[randomIndex] = personAux;
  }

  return arPeople;
}

function getPeopleData()
{
  let spreadsheet = SpreadsheetApp.getActive();

  let sheet = spreadsheet.getActiveSheet();

  let range = sheet.getRange(2, 1, 20, 3).activate();

  return range.getValues();
}

function compareWeight(a, b) 
{
  if(a.weight < b.weight)
    return -1;

  if(a.weight > b.weight)
    return 1;

  return 0;
}

function buildPayload(facilitatorName) {
  let payload = {
    "blocks": [
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": ":bell: *Pessoa Facilitadora Da Daily* :bell:"
        }
      },
      {
        "type": "divider"
      },
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": "Olá,"
        }
      },
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": "Após sorteio, a daily será facilitada pela pessoa: " + facilitatorName + "."
        }
      },
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": "Um bom dia e bom trabalho."
        }
      },
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": "A pessoa não está disponível?"
        },
        "accessory": {
          "type": "button",
          "text": {
            "type": "plain_text",
            "text": "Sortear Novamente",
            "emoji": true
          },
          "value": "click_me_123",
          "action_id": "button-action",
          "accessibility_label": "Sortear Novamente",
          "url": webappURL
        }
      }
    ]
  };

  return payload;
}

function sendAlert(payload) {
  if (webhookURL != "")
  {
    var options = 
    {
      "method": "post", 
      "contentType": "application/json", 
      "muteHttpExceptions": true, 
      "payload": JSON.stringify(payload) 
    };
    
    try 
    {
      UrlFetchApp.fetch(webhookURL, options);
    } 
    catch(e) 
    {
      Logger.log(e);
    }
  }
  else
    Logger.log("Integração com Slack não ocorreu. URL do Webhook está vazia.")
}

function updateLastExecution()
{
  let spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('E2').activate();
  spreadsheet.getCurrentCell().setValue(new Date());
  spreadsheet.getRange('E3').activate();
}
