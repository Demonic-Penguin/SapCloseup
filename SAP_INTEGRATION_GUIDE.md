# SAP GUI Integration Guide

This guide explains how to connect the SAP Close-Up Automation web application to a real SAP GUI system running on Windows.

## Overview

The integration works by:
1. Running a small API adapter server on the Windows machine with SAP GUI
2. Configuring the web application to communicate with this adapter
3. The adapter translates API calls into SAP GUI scripting actions

## Requirements

### On the Windows Machine:
- Windows 10 or later
- SAP GUI installed (with scripting enabled)
- Python 3.6+ installed
- Access to SAP system

### For the Web Application:
- Access to configure environment variables
- Network connectivity to the Windows machine

## Step 1: Set Up the Windows Adapter

1. Copy these files to your Windows machine with SAP GUI:
   - `sap_api_adapter.py`
   - `start_sap_adapter.bat`

2. Install required Python packages:
   ```
   pip install flask requests
   ```

3. Install the Windows-specific package:
   ```
   pip install pywin32
   ```

4. Edit the `start_sap_adapter.bat` file:
   - Change the `SAP_API_KEY` to a secure value
   - Change the `HOST` if needed (default is `0.0.0.0` which allows all network connections)
   - Change the `PORT` if needed (default is `5001`)

## Step 2: Enable SAP GUI Scripting

1. Open SAP GUI
2. Go to Customizing of Local Layout
3. Check the "Enable Scripting" option
4. Restart SAP GUI

Note: This might require administrator assistance if your SAP environment has restricted permissions.

## Step 3: Configure the Web Application

Set the following environment variables for the web application:

```
SAP_CONNECTION_TYPE=api
SAP_API_URL=http://<windows_machine_ip>:5001/api/sap
SAP_API_KEY=<same_key_as_in_batch_file>
```

Replace `<windows_machine_ip>` with the IP address of your Windows machine running SAP GUI.

## Step 4: Start the Integration

1. On the Windows machine:
   - Make sure SAP GUI is open and logged in
   - Run the `start_sap_adapter.bat` file

2. The adapter should start and display:
   ```
   Starting SAP API Adapter on 0.0.0.0:5001
   API Key is set to: your-key-here
   Make sure SAP GUI is running and accessible
   ```

3. Leave this command window open while using the application

## Step 5: Test the Connection

1. Open the web application
2. Enter a service order number
3. The application should connect to SAP through the adapter and retrieve real data

## Troubleshooting

### Connection Errors
- Make sure the Windows firewall allows connections on the specified port
- Verify the IP address is correct in the SAP_API_URL
- Ensure the API keys match between adapter and web app

### SAP GUI Errors
- Verify SAP GUI is open and logged in
- Check that scripting is enabled in SAP GUI
- Look at the adapter console for error messages
- Check the `sap_adapter.log` file for detailed errors

### Permission Errors
- Make sure you have the necessary SAP permissions
- Check if running the adapter as administrator helps

## Security Considerations

- The API adapter currently uses a simple API key for authentication
- Ensure your network configuration restricts access to the adapter
- Consider implementing additional security measures in a production environment
- Never expose the adapter to the public internet without proper security

## Common Issues and Solutions

### SAP GUI Screen Changes
If SAP updates or changes its interface, the adapter may need to be updated to match the new screen elements.

### Performance
If you notice slow performance, it may be due to:
- Network latency between web app and adapter
- SAP GUI performance issues
- High system load on the Windows machine

### API Adapter Crashes
If the adapter crashes:
1. Check the log file for errors
2. Ensure SAP GUI is stable and not showing error dialogs
3. Restart both SAP GUI and the adapter

## Support

For technical support, contact the IT support team.

For issues with the SAP Close-Up process, contact the operations team.