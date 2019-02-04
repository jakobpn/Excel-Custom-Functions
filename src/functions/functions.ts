const endPoint = "https://gateway.saxobank.com/sim";
const accessToken =
"eyJhbGciOiJFUzI1NiIsIng1dCI6IkQ2QzA2MDAwMDcxNENDQTI5QkYxQTUyMzhDRUY1NkNENjRBMzExMTcifQ.eyJvYWEiOiI3Nzc3NyIsImlzcyI6Im9hIiwiYWlkIjoiMTA5IiwidWlkIjoiT3ptMWk5M0QtLWdHdEFTYzRURGpxUT09IiwiY2lkIjoiT3ptMWk5M0QtLWdHdEFTYzRURGpxUT09IiwiaXNhIjoiRmFsc2UiLCJ0aWQiOiIyMDAyIiwic2lkIjoiNTAyN2I4ZTQ3ZmI4NDAwZWE4YzMzZDQzMjExZGMyMzYiLCJkZ2kiOiI4NCIsImV4cCI6IjE1NDkzNDE5NDQifQ.3JXChxvGjlFE535mlZ5V0XfCevhZMtpFY_--UtPTIsQbi8-eRrE3ZrNV0plMmw7eisvSjns1Iw1KTgAyI39v_g";

function openApiGet(uri: string, parameterList: string): Promise<any[][]>
{
  return new Promise(function (resolve) {
    var xhr = new XMLHttpRequest();
    
    var url = `${endPoint}${uri}`;
    console.log(`OpenAPI Request: ${url}`);

    //add handler for xhr
    xhr.onreadystatechange = function () {
      if (xhr.readyState == XMLHttpRequest.DONE) {
        //return result back to Excel
        const response = JSON.parse(xhr.responseText);
        console.log(`OpenAPI: ${JSON.stringify(response)}`);
        const result = formatResult(response, parameterList);
        console.log(`Result: ${JSON.stringify(result)}`);
        resolve(result);
      }
    };

    //make request
    xhr.open("GET", url, true);
    xhr.setRequestHeader("Authorization",`Bearer ${accessToken}`);
    xhr.send();
  });
}

function formatResult(data: object, parametersList: string): any[][] {
  const parameters = parametersList.split(",");
  const header = parameters.map((p) => p.trim().replace(/.*\./, ""));
  const body = data["Data"].map((item) => parameters.map((p) => eval(`item.${p}`)));

  return [header].concat(body);
}

function stockQuote(ticker: string): Promise<number> {
  //1. If(web call), you should want to return a "Promise".
  //2. This tells Excel to #GETTING_DATA, until the promise is 'resolved'.
  //3. NEW: Custom Functions allow you the user to continue interaction
  return new Promise(function(resolve) {
    var xhr = new XMLHttpRequest();

    var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";

    //add handler for xhr
    xhr.onreadystatechange = function() {
      if (xhr.readyState == XMLHttpRequest.DONE) {
        //return result back to Excel
        var price = parseFloat(xhr.responseText);

        resolve(price);
      }
    };

    //make request
    xhr.open("GET", url, true);
    xhr.send();
  });
}


/**
 * Adds two numbers.
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
function add(first: number, second: number): number {
  return first + second;
}

/**
 * Displays the current time once a second.
 * @param handler Custom function handler  
 */
function clock(handler: CustomFunctions.StreamingHandler<string>): void {
  const timer = setInterval(() => {
    const time = currentTime();
    handler.setResult(time);
  }, 1000);

  handler.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time.
 * @returns String with the current time formatted for the current locale.
 */
function currentTime(): string {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @param incrementBy Amount to increment
 * @param handler Custom function handler 
 */
function increment(incrementBy: number, handler: CustomFunctions.StreamingHandler<number>): void {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    handler.setResult(result);
  }, 1000);

  handler.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @param message String to write.
 * @returns String to write.
 */
function logMessage(message: string): string {
  console.log(message);

  return message;
}

/**
 * Defines the implementation of the custom functions
 * for the function id defined in the metadata file (functions.json).
 */
CustomFunctions.associate("GET", openApiGet);
CustomFunctions.associate("STOCKQUOTE", stockQuote);
CustomFunctions.associate("ADD", add);
CustomFunctions.associate("CLOCK", clock);
CustomFunctions.associate("INCREMENT", increment);
CustomFunctions.associate("LOG", logMessage);
