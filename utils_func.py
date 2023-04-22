import pandas as pd
import numpy as np
from matplotlib import pyplot as plt
import seaborn as sns
from tqdm.notebook import tqdm, trange

import os, sys, re, warnings, glob, functools
from os.path import expanduser
from pathlib import Path
from datetime import datetime

from sklearn.model_selection import train_test_split, StratifiedKFold

# ROC visualization
from scipy import interp
from itertools import cycle
from sklearn import svm, datasets
from sklearn.metrics import roc_curve, auc
from sklearn.preprocessing import label_binarize, OrdinalEncoder, StandardScaler, MinMaxScaler
from sklearn.multiclass import OneVsRestClassifier
from sklearn.metrics import log_loss, roc_auc_score
from sklearn.linear_model import LogisticRegression
from sklearn.model_selection import GridSearchCV

from glob import glob

import shutil

tqdm.pandas()
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

