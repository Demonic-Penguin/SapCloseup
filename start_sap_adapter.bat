@echo off
echo Setting up SAP API Adapter...

REM Set environment variables
set PORT=5001
set HOST=0.0.0.0
set SAP_API_KEY=change-this-in-production

echo Starting SAP API Adapter on %HOST%:%PORT%
echo Make sure SAP GUI is running and logged in before starting the adapter

REM Start the adapter
python sap_api_adapter.py

pause