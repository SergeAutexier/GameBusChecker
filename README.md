# GameBusChecker

A tool for checking the campaigns exported from Gamebus in xlsx format.

## Installation and use

1. **Install Python** 
    Tested on Python 3.13, but should work for 3.10 to 3.13

    Download and install from 
    ```
    https://www.python.org/downloads/
    ```


2. **Set up Python environment (Windows)**: 

  On the command line, create a directory and change to that directory. Then:

   ```
   python -m venv .venv
   .venv\Scripts\activate
   pip install -r requirements.txt
   ```

   **Set up Python environment (Linux/MacOSX)**

   On the command line, create a directory and change to that directory. Then:
  
   ```
   python -m venv .venv
   source .venv/Scripts/activate
   pip install -r requirements.txt
   ```

3. **Run**:

  Download a Gamebus campaign from your campaign at (menu in the 3 dots next to your campaign)

  ```
  https://campaigns.healthyw8.gamebus.eu/editor/campaigns
  ```

  and save the Excel file CAMPAIGN_FILENAME.xlsx in a directory, e.g. PATH. 

   ```
   python Gamebus_Campaign_Checker.py -f PATH/CAMPAIGN_FILENAME.xlsx # (Linux/MacOSX) Run all checks on the campaign downloaded from Gamebus
   python Gamebus_Campaign_Checker.py -f PATH\CAMPAIGN_FILENAME.xlsx # (Windows) Run all checks on the campaign downloaded from Gamebus
   python Gamebus_Campaign_Checker.py -h                             # see how to run individual checks only
   ```

  
