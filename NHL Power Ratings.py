#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np
import datetime
from sklearn.metrics import accuracy_score
from sklearn.linear_model import Ridge
from openpyxl import load_workbook


# In[2]:


#load the data
df = pd.read_excel("NHL Scores.xlsx", skiprows=1, names=["date", "visitor","visitor_goals", "home", "home_goals"])
df["date"] = pd.to_datetime(df["date"],format="%Y-%m-%d")

#Only stats up to yesterdays completed games

today = pd.Timestamp("today").floor("D")

df = df.loc[(df.date < today)]


#Add new feature "Goal Difference"
df["goal_difference"] = df["home_goals"] - df["visitor_goals"]


# In[3]:


#New variables to show home team win or loss result
df["home_win"] = np.where(df["goal_difference"] > 0, 1, 0)
df["home_loss"] = np.where(df["goal_difference"] < 0, 1, 0)


# In[4]:


#Create dummy variable matrices for df_visitor and df_home
df_visitor = pd.get_dummies(df["visitor"], dtype=np.int64)
df_home = pd.get_dummies(df["home"], dtype=np.int64)


# In[5]:


#Combine previous results from df_visitor and df_home to get final dataset
#Subtract df_visitor from df_home to get df_model
df_model = df_home.sub(df_visitor)
df_model["goal_difference"] = df["goal_difference"]


# In[11]:


#Removing any NaN values from dataset
df_model = df_model.fillna(0)


# In[7]:


df_train = df_model

lr= Ridge(alpha = 0.001)
X = df_train.drop(["goal_difference"], axis = 1)
y= df_train["goal_difference"]

lr.fit(X,y)


# In[8]:


df_ratings = pd.DataFrame(data={'team': X.columns, 'rating': lr.coef_,'home_adv':lr.intercept_})


# In[9]:


df_ratings["rank"] = df_ratings["rating"].rank(ascending=False)
df_ratings = df_ratings.sort_values('rank')
df_ratings


# In[10]:


def write_excel(filename,sheetname,dataframe):
    with pd.ExcelWriter(filename, engine="openpyxl", mode = "a",index=False) as writer:
        workBook = writer.book
        try:
            workBook.remove(workBook[sheetname])
        except:
            print("Worksheet does not exist")
        finally:
            dataframe.to_excel(writer, sheet_name=sheetname)
            writer.save()
write_excel('NHL Power Ratings.xlsx', 'Power Ratings', df_ratings)


# In[ ]:




