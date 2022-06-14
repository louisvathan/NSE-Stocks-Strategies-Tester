# NSE-Stocks-Strategies-Tester
Takes EOD data from Yahoo Finance and runs it through various technical indicators and gives the output in Excel Format.

## Getting Started
You need the following packages to run this program.
* Pandas == 1.2.1
* Openpyxl == 3.0.3
* Numpy == 1.19.5
* Matplotlib == 3.2.2
* yfinance == 0.1.70

You have to open Main_V0.1.py to run the program. You can add stock symbols as list in Stocks.py file and link it to the variable "stocks" in line 19 of Main_V0.1.py file. You can also add new set of stocks in Stocks.py file as you desire.

## Contributing
If you wish to give your contributions like fixing bugs, improving things and adding documentation to this project please feel free to give pull request.

## License
This project is licensed under the MIT License - see the LICENSE.md file for details

## Terms & Conditions
Please read these terms and conditions carefully before using Our Service.

## How to Use?
**Step 1:** Download and keep all the files a folder.

**Step 2:** Open Stocks.py in your editor and add a variable with the stocks you desire as a list or edit the existing variables. Then save the file.

```
mystocks = ['HDFCBANK', 'SBIN']
```

**Step 3:** Open Main_V0.1.py in your editor and update the "stocks" variable (at Line 3) as per your desired variable from the Stocks.py file. Then save the file.

```
# USER INPUT
import Stocks
stocks = Stocks.mystocks    # Edit this line

```

**Step 4:** Run the Main_V0.1.py file in your command line or editor. The program will take some time to execute completely.

## Interpreting the Output
Once the program excuted successfully, you can find a new folder is created in the same location of the program files with a naming syntax as "Results (YYYYMMDD)". Open that folder and you find the following folder structure.

![File Structure](https://github.com/louisvathan/NSE-Stocks-Strategies-Tester/blob/main/Sample/Images/Folder%20Structure.jpg?raw=true)
