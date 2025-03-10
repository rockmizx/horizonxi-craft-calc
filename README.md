# Horizon XI Item Price Calculator

This is a Python script that calculates the production cost of items in the game Horizon XI, based on the current prices of the required materials. It uses Selenium to automate price collection directly from the game's website and generates an Excel file with the results.

## Dependencies

To run this script, you will need the following Python libraries:

- `tkinter` (comes pre-installed with Python)
- `selenium`
- `openpyxl`
- `regex`

You can install the dependencies using `pip`:

```bash
pip install selenium openpyxl regex
```

## Geckodriver Installation

The script uses `geckodriver` to control the Firefox browser. Follow the steps below to set up `geckodriver`:

1. **Download Geckodriver**: Visit the [geckodriver releases page](https://github.com/mozilla/geckodriver/releases) and download the version corresponding to your operating system.

2. **Extract Geckodriver**: Extract the downloaded file and place the `geckodriver` executable (or `geckodriver.exe` on Windows) in the same folder as the `horizon_item_calc.py` script.

3. **Verify the Path**: Ensure the path to `geckodriver` is correct in the script. The script expects `geckodriver` to be in the same folder as itself.

## Firefox Configuration

The script uses an existing Firefox profile to access the Horizon XI website. You will need to configure the path to the Firefox executable and the profile you want to use.

1. **Firefox Path**: In the `config.json` file, define the path to the Firefox executable (`firefox_binary_path`). For example:

    ```json
    {
        "firefox_binary_path": "C:/Program Files/Mozilla Firefox/Firefox.exe",
        "firefox_profile_path": "C:/Users/XXXXX/AppData/Roaming/Mozilla/Firefox/Profiles/XXXXX.default-release"
    }
    ```

2. **Profile Path**: Define the path to the Firefox profile you want to use (`firefox_profile_path`). You can find the profile path in the `Profiles` folder within the Firefox installation directory.

## How to Use

1. **Run the Script**: Execute the `horizon_item_calc.py` script using Python:

    ```bash
    python horizon_item_calc.py
    ```

![Captura de tela 2025-03-10 002622](https://github.com/user-attachments/assets/bdd39ffc-7758-4c38-80ad-9e3e7dca8e53)

2. **Graphical Interface**: A window will open with a simple graphical interface.

3. **Start Browser**: Click the "Start Browser" button to open Firefox and access the Horizon XI website.

4. **Select Item**: From the dropdown menu, select the item you want to calculate.

5. **Calculate**: Click the "Calculate" button to start the price collection and cost calculation process.

6. **Results**: The script will collect material prices, calculate the total production cost, and display the results in a message box. Additionally, an Excel file named `ff11_crafts.xlsx` will be generated with the calculation details.

![Captura de tela 2025-03-10 002748](https://github.com/user-attachments/assets/4b5c2ce0-3300-4a13-830f-c9b7ed7f2333)

## Example Usage

1. **Start the Script**: Run the script and click "Start Browser".
2. **Select Item**: Choose "Meat Mithkabob" from the dropdown menu.
3. **Calculate**: Click "Calculate" to see the production cost of "Meat Mithkabob".
4. **Results**: The script will display the total material cost, the item's selling price, and the expected profit.

## Configuration Files

- **items.json**: Contains the list of items and their required materials for production. You can add or modify items in this file.
- **config.json**: Contains the path to the Firefox executable and the profile to be used by the script.

## Notes

- Ensure Firefox is installed and the configured profile is logged into the Horizon XI website.
- The script may take a few seconds to collect item prices, depending on your internet connection speed.
- The generated Excel file (`ff11_crafts.xlsx`) can be opened with any Excel-compatible program for a more detailed cost analysis.

## Contributions

Feel free to contribute improvements or fixes to the script. Simply fork the repository and submit a pull request.
