# Patent Finder

Patent Finder is a program that searches for patents on Google Patents and extracts the following information from them:
- Patent number
- Google Patent Link
- Patent Title
- Filing Date
- Expiry Date
- Abstract
- Claim 1*
- Patent Type*
- Main Topic/Subject of the invention*
- Patent Review*

\* These headers are derived with AI. Thus, they might not be 100% accurate

This information is then put into an excel file.

## Installation
To run this project on your computer, first navigate to the folder you wish to store it in and clone it
```
git clone 
```

After cloning the project, please install the dependencies by running this command:
```bash
pip install -r requirements.txt
```

## Usage

1. Create a folder within the directory called "Files". This is where the exported file will go.

2. Go to [https://aistudio.google.com/apikey](https://aistudio.google.com/apikey) and request a free Gemini API key. Do not share this with anyone else!

3. Create a file within the directory called ".env" and paste your API key like so:
```
GEMINI_KEY= YOUR API KEY HERE
```

The folder structure of this project can be represented as the following
```
Patent Finder
|   README.md
|   .env
|   gui.py
|   main.py
|   script.py
|   requirements.txt
└───Files
    └───Saved exports will be found here
```

4. The input file must be an Excel file with the header "Patent Number" (Cell A1). The patent numbers you wish to extract information from should go below this header. 

5. To run the program, navigate to your directory and run the following command:
```bash
python main.py
```
