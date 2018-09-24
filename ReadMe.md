
# Microsoft Teams Graph APIs - Application Permission

This is a console sample which shows how to create a new Team in Microsoft Teams using Graph APIs.

### How to see the Sample code running
1. Application Registration 
    1. [Register a new Application](https://developer.microsoft.com/en-us/graph/docs/concepts/auth_v2_service).
    1. Generate a New Password and save it.
    1. Add web platform with Redirect URLs as `http://localhost/myapp/permissions`.
    1. Add `Group.ReadWrite.All` & `User.Read.All` in Application Permissions section.
1. Update `tenant`, `appId` & `appSecret` values.
1. For the first time, keep `GetOneTimeAdminConsent()` method uncomment. This will open up browser window and ask for admin consent.

>**Note**: This sample code needs one time admin consent.

## More Information
For more information about getting started with Teams, please review the following resources:
- Review [Getting Started with Teams](https://msdn.microsoft.com/en-us/microsoft-teams/setup)
- Review [Getting Started with Bot Framework](https://docs.microsoft.com/en-us/bot-framework/bot-builder-overview-getstarted)
- Review [Testing your bot with Teams](https://msdn.microsoft.com/en-us/microsoft-teams/botsadd)

