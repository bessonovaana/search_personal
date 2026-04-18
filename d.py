import matplotlib.pyplot as plt
import pandas as pd
import os, pathlib

dirs = pathlib.Path('')
struct = ['.csv', '.json', '.parquet']
doc = ['.pdf', '.doc', '.docx', '.rtf', '.xls']
web = ['.html']
image = ['.tif', '.jpeg', '.png', '.gif']
video = ['.mp4']

files = [()]