import streamlit as st
import pandas as pd
import sys
import smtplib
import logging
import numpy as np
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from python_topsis_Samridhi_101916086 import *
from python_topsis_Samridhi_101916086 import topsis as tp
def send_mail(email_id,resultfile):
    fromaddr = "sgarg7_be19@thapar.edu"
    toaddr = email_id
   
# instance of MIMEMultipart
    msg = MIMEMultipart()
  
# storing the senders email address  
    msg['From'] = fromaddr
  
# storing the receivers email address 
    msg['To'] = toaddr
  
# storing the subject 
    msg['Subject'] = "TOPSIS SCORE AND RANK GENERATOR" 
  
# string to store the body of the mail
    body = '''For the given input file(.csv), here is your ouput(.csv) file with topsis score and rank information provided for MCDM(Multiple Criteria Decision Making'''
  
# attach the body with the msg instance
    msg.attach(MIMEText(body, 'plain'))
  
# open the file to be sent 
    filename = resultfile
    attachment = open(resultfile, "rb")
  
# instance of MIMEBase and named as p
    p = MIMEBase('application', 'octet-stream')
  
# To change the payload into encoded form
    p.set_payload((attachment).read())
  
# encode into base64
    encoders.encode_base64(p)
   
    p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
  
# attach the instance 'p' to instance 'msg'
    msg.attach(p)
  
# creates SMTP session
    s = smtplib.SMTP('smtp.gmail.com', 587)
  
# start TLS for security
    s.starttls()
  
# Authentication
    s.login(fromaddr, "Samridhi3404##")
  
# Converts the Multipart msg into a string
    text = msg.as_string()
  
# sending the mail
    s.sendmail(fromaddr, toaddr, text)
  
# terminating the session
    s.quit()
def inDigit(x):
  try:
    return float(x)
  except ValueError:
        logging.error("Weights: Enter numeric values!!")
        # st.error("Weights: Enter numeric values!!")
        raise Exception("Weights: Enter numeric values!!")
        logging.shutdown()    
def topsis(df,weights,impacts,resultfile):
    logging.basicConfig(filename="101916086-log.txt", level=logging.INFO)
    try:
        
        for col in df.columns:
            if df[col].isnull().values.any():
                logging.error(f"{col} contains null values")
                # st.error(f"{col} contains null values")
                raise Exception(f"{col} contains null values")      
                logging.shutdown() 

        if len(df.columns)<3:
            logging.error("No of inappropriate columns, minimum 3 needed")
            # st.error("No of inappropriate columns, minimum 3 needed")
            raise Exception("No of inappropriate columns, minimum 3 needed")      
            logging.shutdown() 
        i=0
        weights = weights.split(",")
        impacts = impacts.split(",")
        data=df.iloc[:,1:]    
        for p in range(len(weights)):
            weights[p] = inDigit(weights[p])
        if data.shape[1] != data.select_dtypes(include=["float", 'int']).shape[1]:
            logging.error("Columns numst contain only numeric values")
            # st.error("Columns numst contain only numeric values")
            raise Exception("Columns numst contain only numeric values")
            logging.shutdown()

        if(len(weights)!=len(data.columns)):
            logging.error("weights are less/more in number!!")
            # st.error("weights are less in number!!")
            raise Exception("weights are less in number!!")
            logging.shutdown()
        if(len(impacts)!=len(data.columns)):
            logging.error("Impacts are less/more in number!!")
            # st.error("Impacts are less in number!!")
            raise Exception("Impacts are less/more in number!!")
            logging.shutdown()
        best_val = []
        worst_val = []
        for column in data.columns:
            sq = data[column].pow(2).sum()
            sq = np.sqrt(sq)
            data[column] = data[column]*weights[i]/sq
            if(impacts[i]=='+'):
                best_val.append(data[column].max())
                worst_val.append(data[column].min())
            elif(impacts[i]=='-'):
                best_val.append(data[column].min())
                worst_val.append(data[column].max())
            else:
                logging.error("Impacts: Invalid input enter only '+' or '-'")
                # st.error("Impacts: Invalid input enter only '+' or '-'")
                raise Exception("Impacts: Invalid input enter only '+' or '-'")
                logging.shutdown()
            i+=1
        euclid_best=0
        euclid_worst=0
        topsis_score=[]
        column = len(data.columns)
        row = len(data)
        for x in range(row):
            for y in range(column):
                euclid_best+=(data.iloc[x][y]-best_val[y])**2
                euclid_worst+=(data.iloc[x][y]-worst_val[y])**2
            euclid_worst = np.sqrt(euclid_worst)
            euclid_best = np.sqrt(euclid_best) + euclid_worst
            topsis_score.append(euclid_worst/euclid_best)
        data["Topsis_score"]=topsis_score  
        topsis_score = pd.DataFrame(topsis_score)
        topsis_rank = topsis_score.rank(method='first',ascending=False)
        data["Rank"]=topsis_rank
        # data.insert(column+2,"Rank",topsis_score,allow_duplicates=False)
        # print(dataset)
        data.to_csv(resultfile,index=False)  
       
        send_mail(email_id,resultfile)
        st.success("Check your email, result file is successfully sent")
        
    except IOError:
        logging.error("file not found!!")
        # st.error("file not found!!")
        raise Exception("file not found!!")
        logging.shutdown()  
# def save_uploaded_file(file):
#     with open(file, "wb") as f:
#         f.write(buf.getbuffer())

if __name__ == '__main__':    
    st.title("Multiple Criteria Decision Making using Topsis")
    st.write("""
    ## Topsis Score and rank generator
    """)
    st.write("""
    #### To get the result: pip install python-topsis-Samridhi-101916086==4.1.1
    """)
    try:
        spectra = st.file_uploader("upload file", type={"csv", "xlsx"})
        # print(spectra.name)
        if spectra is not None:
            
            # spectra_df.to_csv('file_new')
            # print(type(spectra_df))
            st.success("Successfully uploaded input file")   
            # st.write(spectra_df)
            weights = st.text_input("Weights")
            impact = st.text_input("Impacts")
            email_id = st.text_input("Email")
            resultfile = st.text_input("Resultfile")
            # st.write(resultfile)
            submit_button=st.button("Send")
            if submit_button:
                try:
                    if email_id.split('@')[1]!="thapar.edu":
                        st.error("Invalid Email supplied, provide thapar email")  
                    if resultfile.split('.')[1]!="csv" and resultfile.split('.')[1]!="xlsx":
                        st.error("File format must be a csv of excel file")
                    if resultfile.split('.')[1]=="csv":
                        spectra_df = pd.read_csv(spectra)
                    elif resultfile.split('.')[1]=="xlsx":
                        read = pd.read_excel (spectra)
                        read.to_csv ('file_new', index = None,header=True)
                        spectra_df = pd.read_csv(file_new)
                    topsis(spectra_df,weights,impact,resultfile)     
                    
                except Exception as e:
                    st.error(e)
                    # print(e) 
            
            # save_uploaded_file(spectra)  
            
           
    except IOError:
      raise Exception("file not found!!")
 
        

