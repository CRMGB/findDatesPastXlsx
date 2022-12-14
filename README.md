# findDatesPastXlsx
Script required by my friend Tortu just because he forgets when there is a date in the past inside of a xlsx file.

- Get dates from a multiple xlsx files.
- IMPORTANT all xlsx files to check have to be in the files/ directory.
- Output the dates in the past in the A1 cell of each file checked.

###INSTALLATION (We assume you've got python 3.7 or higher):
  - Check whether you've got pip installed: pip help
  - If there is no pip in your system (For Windows):
      - curl https://bootstrap.pypa.io/get-pip.py -o get-pip.py
      - python get-pip.py
  - pip install -r requirements.txt
  - python slxdates.py
