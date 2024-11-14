## Description
This Python application connects to a personal, work or school Outlook account, downloads the message contents and its attachments, analyses them and generates a consolidated Excel report with the desired columns.

## Requirements
This application requires the following:
1. **Microsoft Entra App registration**: This provides a secured authentication to Outlook mail (the mail where inventory details are received). This app provides a Client ID which can be used to fetch the required tokens to authenticate with the resource server to read the emails.
2. **OpenAI API Key**: This LLM helps in understanding the contents of the email and getting the needed details of the items from the attachment). This API Key helps to access OpenAI model to analyse the email contents.
3. **PriceRunner "Product" token**: This provides details of the registered products to find real time pricing from different merchants in the specified region. A "Product" token acquired from the PriceRunner team helps in fetching product details with the help of GTIN codes.

### Microsoft Entra App Registration

1. Create a new app registration in the [Microsoft Entra platform](https://entra.microsoft.com/)

<img width="1440" alt="1" src="https://github.com/vallabha108/aiexcelgenerator/assets/46571593/2cb6b6d8-f0c8-4dcf-ac2d-1a99d70da1e5">

2. Provide a name for the app and select account type as "Accounts in any organizational directory (Any Microsoft Entra ID tenant - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)"

<img width="1440" alt="2" src="https://github.com/vallabha108/aiexcelgenerator/assets/46571593/85b1a7b1-acb3-48ea-a8f8-c79491d4d7fe">

3. Note down the **Client ID** which you will need while running the app.

<img width="1440" alt="3" src="https://github.com/vallabha108/aiexcelgenerator/assets/46571593/e9d6b81c-604f-4599-97fd-7266048ce135">

4. Navigate to **Authentication** section from the left pane and allow public client flows.

<img width="1440" alt="4" src="https://github.com/vallabha108/aiexcelgenerator/assets/46571593/3b84a7b0-b19d-4632-abf3-d5cd0f79244d">

5.  Navigate to **API Permissions** section from the left pane and click on **Add a permission** option.

<img width="1440" alt="5" src="https://github.com/vallabha108/aiexcelgenerator/assets/46571593/79a740d0-c375-401f-94f5-f4e1ee0c8da3">

6.  Follow the steps in the following screenshots to add two permissions - _**Offline_access**_ and _**Mail.Read**_.

<img width="1440" alt="6" src="https://github.com/vallabha108/aiexcelgenerator/assets/46571593/9c6d7c2e-c763-40ab-8c84-a282d105b778">

<img width="1440" alt="7" src="https://github.com/vallabha108/aiexcelgenerator/assets/46571593/e8b2ea3c-8dcb-49d9-a125-5d434d704ab7">

<img width="1440" alt="8" src="https://github.com/vallabha108/aiexcelgenerator/assets/46571593/2ff312c3-28d3-442a-8f9d-ce6ab7e0b17c">

<img width="1440" alt="9" src="https://github.com/vallabha108/aiexcelgenerator/assets/46571593/3a9ab65a-dfeb-41ca-ad3c-c941cfc9fd60">

Finally, the permissions displayed in the screenshot should exist for this app.

<img width="1440" alt="10" src="https://github.com/vallabha108/aiexcelgenerator/assets/46571593/05a0ba14-5e9f-411d-b6ec-eff4451dbdd7">

### OpenAI API Key creation

1. Create a new API Key with the default permissions from the [OpenAI Platform](https://platform.openai.com/api-keys).

<img width="1440" alt="11" src="https://github.com/vallabha108/aiexcelgenerator/assets/46571593/880b86ce-d7ba-4acc-bea6-7709bd502898">

Note down the **API Secret** which you will need while running the app.

### PriceRunner

Fill the form in the [PriceRunner API page](https://www.pricerunner.com/register/api) to get the test or production level tokens from the PriceRunner team.

## Steps to run the app
1. Clone the repository in your local system.
```
git clone https://github.com/vallabha108/aiexcelgenerator.git
```

```
cd aiexcelgenerator
```
2. (Optional) Switch to `main` branch to test the latest changes.  (for initial clone, this step is not required. This is required only if you are in a different branch).
```
git checkout main
```

3. Install `poetry`, a tool for dependency management and packaging in Python.
```
pipx install poetry
```

> [!TIP]
> If `pipx` is not already installed, you can follow any of the options in the official [pipx installation instructions](https://pipx.pypa.io/stable/installation/).

4. Activate a virtual environment in `poetry`.
```
poetry shell
```

> [!TIP]
> You can exit this virtual environment anytime by typing `exit` and pressing `enter` or `return` key.

5. Install the defined dependencies for the application.
```
poetry install --no-root
```

6. Create `config.yaml` file in the root directory (```aiexcelgenerator```) and store the necessary secrets. You need three secrets to be stored in this case - `openai_api_key` which is the API Secret you got from OpenAI API Key creation, `client_id` which you got from Microsoft Entra App registration and the `price_runner_token` which you got from the PriceRunner team.
```
openai_api_key=<YOUR_OPENAI_API_KEY>
client_id=<YOUR_MICROSOFT_ENTRA_APP_CLIENT_ID>
price_runner_token=<YOUR_PRICE_RUNNER_TOKEN>
```

7. Run the app from the root directory of the application by speciying an *optional* path to a directory where the emails will be downloaded and the report will be saved. Here's an example:
```
python -m src.app --path /Users/admin/Outputs
```

> [!NOTE]
> In this command, I have provided `Users/admin/Outputs` because I need the contents to be saved by the app in that directory.
> If no path is specified, the messages and reports will be saved by default in a directory named `output` in the root directory of the application.
> Here's an example of no path:
> ```
> python -m src.app
> ```

8. Sign-in to Outlook account by navigating to the link and pasting the device code (link and code must be displayed in the CLI by now). Consent the app with the requested read permissions so that the app can read emails and attachments.

<img width="1440" alt="12" src="https://github.com/vallabha108/aiexcelgenerator/assets/46571593/efecb0c0-ba6b-4183-b795-b87024996701">

<img width="1440" alt="13" src="https://github.com/vallabha108/aiexcelgenerator/assets/46571593/d9be3021-fcc1-4736-9fb0-684a2883a60d">

<img width="1440" alt="14" src="https://github.com/vallabha108/aiexcelgenerator/assets/46571593/61c1999b-5ee9-49be-8f02-2228e72251fa">

After consenting, navigate back to the CLI where you can see the live updates of the application process.

![Group 1](https://github.com/vallabha108/aiexcelgenerator/assets/46571593/b7a633d3-ac7c-4342-a128-d526570f8add)

> [!NOTE]
> Logs can be found at the `logs` folder in the root directory.
