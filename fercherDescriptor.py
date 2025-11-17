import requests
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# Configuarion
PAT = "ghp_yourpersonalaccesstokenhere"
organizations = ["Org1", "Org2"]

session = requests.Session()
session.auth = ("", PAT)

