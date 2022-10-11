from sqlalchemy import create_engine
import requests
import time
import pandas as pd
from datetime import datetime
import numpy as np
import os

pd.set_option("display.max_columns", None)

HOST = "MERU_DATABASE@gua73543.snowflakecomputing.com"
engine = create_engine(
    "snowflake://{user}:{password}@{account_identifier}/MERU_DATABASE?warehouse=COMPUTE_WH".format(
        user="jose_alvarez",
        password="jwe*nxp8uhz3VBC.jqa",
        account_identifier="gua73543",
    )
)


headers = {
    "Authorization": "Bearer keyVa6CKp4p6D8oM9",
    "Content-Type": "application/json",
}
