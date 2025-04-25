import os
from dotenv import load_dotenv

load_dotenv()

f = os.getenv('DATAVERSE_ENTITY_FILTER_COLUMN')
filt = "".join([f, " eq '1234567890'"])
print(filt)