// This file is auto generated by Teams Toolkit to provide you instructions and reference code to call your API.

/*
Next steps:
1. Run npm install. We added the @microsoft/teamsfx to your package.json.
   You need to run the command under the "bot" folder (instead of your project root folder).
2. Refer to the sample code and comments in this file to implement your custom auth provider.

You can import the API client (an Axios instance) in another file and call getPriorityMessages APIs and authentication is now handled for you automatically.

Here is an example for a GET request to "relative_path_of_target_api":
```
const { getPriorityMessagesClient } = require("relative_path_to_this_file");

const response = await getPriorityMessagesClient.get("relative_path_of_target_api");
// You only need to enter the relative path for your API.
// For example, if you want to call api https://my-api-endpoint/test and you configured https://my-api-endpoint as the API endpoint,
// your code will be: const response = await getPriorityMessagesClient.get("test");

const responseBody = response.data;
```

If you added this API while local debugging, stop local debugging and start again because local debugging will not hot reload changes to `.env.teamsfx.local`.

Refer to https://aka.ms/teamsfx-connect-api to learn more. 
*/
const teamsfxSdk = require("@microsoft/teamsfx");

// A custom authProvider implements the `AuthProvider` interface.
// This sample authProvider implementation will set a custom property in the request header
class CustomAuthProvider {
  customProperty;
  customValue;

  constructor(customProperty, customValue) {
    this.customProperty = customProperty;
    this.customValue = customValue;
  }

  // Replace the sample code with your own logic.
  AddAuthenticationInfo = async (config) => {
    if (!config.headers) {
      config.headers = {};
    }
    config.headers[this.customProperty] = this.customValue;
    return config;
  };
}

// Load application configuration
const teamsFx = new teamsfxSdk.TeamsFx();

const authProvider = new CustomAuthProvider(
  // You can also add configuration to the file `.env.teamsfx.local` and use `TeamsFx.getConfig("{setting_name}")` to read the configuration. For example:
  //  teamsFx.getConfig("TEAMSFX_API_GETPRIORITYMESSAGES_CUSTOM_PROPERTY"),
  //  teamsFx.getConfig("TEAMSFX_API_GETPRIORITYMESSAGES_CUSTOM_VALUE")
  "customPropery",
  "customValue"
);
// Initialize a new axios instance to call getPriorityMessages
const getPriorityMessagesClient = teamsfxSdk.createApiClient(
  teamsFx.getConfig("TEAMSFX_API_GETPRIORITYMESSAGES_ENDPOINT"),
  authProvider
);
module.exports.getPriorityMessagesClient = getPriorityMessagesClient;

/* 
Setting API configuration for cloud environment: 
We have set configuration in `.env.teamsfx.local` based on your answers. 
Before you deploy your code to Azure using TeamsFx, follow https://aka.ms/teamsfx-add-appsettings to add the following configuration (with their appropriate values) to your Azure environment: 
TEAMSFX_API_GETPRIORITYMESSAGES_ENDPOINT

Refer to https://aka.ms/teamsfx-connect-api to learn more. 
*/