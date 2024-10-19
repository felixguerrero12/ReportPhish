# Deploying the PhishingRibbon Add-in to Outlook

Deploying the PhishingRibbon add-in to Outlook involves several steps. Here's a general guide on how to deploy a VSTO (Visual Studio Tools for Office) add-in like yours:

## 1. Build the Solution

- In Visual Studio, set the solution configuration to "Release"
- Build the solution (Build > Build Solution)

## 2. Create a deployment package

- Right-click on your project in Solution Explorer
- Select "Publish"
- Choose a publish method (e.g., File System, FTP, Web Site)
- Set a location for the published files
- Click "Publish"

## 3. Distribute the files

- The publish process will create a setup.exe file and a .vsto file
- These files, along with any other dependencies, need to be distributed to users

## 4. Installation on user machines

- Users run the setup.exe file to install the add-in
- Outlook must be closed during installation

## 5. Group Policy Deployment (for organizations)

- Create a network share with the installation files
- Use Group Policy to deploy the add-in across the organization

## 6. ClickOnce Deployment

- This allows users to install from a web server and receive automatic updates
- Publish to a web server
- Users navigate to the .vsto file URL to install

## 7. Manage Trust

- Ensure the add-in is signed with a trusted certificate
- Configure Outlook's Trust Center settings to allow the add-in

## 8. Centralized Deployment (for Microsoft 365)

- Use the Microsoft 365 Admin Center to deploy add-ins centrally

## For your specific add-in

1. Ensure all necessary files are included in the deployment package
2. Update the security email address from the demo address to the actual one
3. Test the deployment in a controlled environment before wide distribution

**Remember:** The exact steps might vary depending on your organization's IT policies and infrastructure. Always consult with your IT department for the best deployment strategy in your specific environment.
