import os
from datetime import datetime
import datetime as dtm
import tarfile
import shutil
import gzip
from exchangelib import Credentials, Account, DELEGATE, HTMLBody,FileAttachment
from exchangelib import Message, Mailbox,Configuration
import pandas as pd
date =  datetime.now().strftime('%d-%m-%Y')
base_dir_format = datetime.now().strftime('%d.%m.%y')
base_dir = f's:\\{base_dir_format}'



def insert_or_update(df, row_id, new_value):
    if df['date'].eq(row_id).any():  # Check if row exists
        df.loc[df['date'] == row_id, 'active'] = new_value  # Update the value
    else:
        new_row = pd.DataFrame({'date': [row_id], 'active': [new_value]})  # Create new row
        df = pd.concat([df, new_row], ignore_index=True)  # Append new row
    
    return df


def compress_files_gzip(file_list, output_filename, compression_ratio):
    with open(output_filename, 'wb') as output_file:
        with gzip.GzipFile(fileobj=output_file, compresslevel=compression_ratio) as gzip_file:
            for file_name in file_list:
                with open(file_name, 'rb') as input_file:
                    shutil.copyfileobj(input_file, gzip_file)


def write_log(message):
    try:
        date =  datetime.now().strftime('%d-%m-%Y')
        with open(date + ' - log.txt', 'a') as file:
            current_date_time = datetime.now().strftime('[%d.%m.%y-%H:%M:%S] ')
            file.write(current_date_time + message + '\n')
            return 'OK'
        
    except Exception as e:
        return e


def addInitials(r):
    if (('_' in r['Ftnr'])):
        return r['Ftnr']
    else:
        if(r['Doc Type_x']=='Invoice'):
            return 'INV_'+r['Ftnr']
        
        if(r['Doc Type_x']=='Payment'):
            return 'PMT_'+r['Ftnr']
        
        if(r['Doc Type_x']=='JV'):
            return 'JV_'+r['Ftnr']
        
        if(r['Doc Type_x']=='Debit Note'):
            return 'DN_'+r['Ftnr']
        
        if(r['Doc Type_x']=='Credit Note'):
            return 'CN_'+r['Ftnr']
        


def create_folder_if_not_exists(folder_path):
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)


def compress_files(output_filename,directory,listOfFiles):
    # List all files in the directory
    
    # Create a tar.gz archive using tarfile with custom compression options
    with tarfile.open(output_filename, "w:gz", compresslevel=9) as tar:
        for file in listOfFiles:
            file_path = os.path.join(directory, file)
            tar.add(file_path)



def search_files(folder_path, file_name):
    # write_log(f'[Search Method] {folder_path} , {file_name}')
    found_files = []
    for foldername, subfolders, filenames in os.walk(folder_path):
        for filename in filenames:
            # write_log(f'[Search] {filename}')
            if file_name.lower() in filename.lower():
                found_files.append(os.path.join(foldername, filename))

    return found_files


def list_files(directory):
    return os.listdir(directory)


def checkModificationDate(file):
    creation_time = os.path.getmtime(file)
    creation_date = dtm.datetime.fromtimestamp(creation_time)
    today = dtm.datetime.now().date()
    return creation_date.date() == today


def sendMail(subject,body,to='mturky@cne.com.eg',cc=None,attachments=None,username='mturky',password='m0h@mmed'):
    
    credentials = Credentials(username=username, password=password)
    config = Configuration(server="mail.cne.com.eg", credentials=credentials)
    myaccount = Account(primary_smtp_address='mturky@cne.com.eg',credentials = credentials, config=config)

    m = Message(
        account=myaccount,
        subject=f'[Auto] {subject}',
        body=HTMLBody( body),
        to_recipients=[to],
        cc_recipients=cc,
        bcc_recipients=[]
    )

    if attachments:

        for a in attachments:
            att = FileAttachment(
            name=f'{a}',
            content_type="image/png",
            is_inline=False,
            is_contact_photo=False,
            content=open(f'{a}', "rb").read())
            m.attach(att)
    m.send()





def getModificationDate(file):
    return os.path.getmtime(file)


def createDTHfile(beinfile,dth_file_name):
        
        df = pd.read_csv(beinfile,dtype='str')

        
        #------------------------ create DTH file

        dth_sub_types = ['beIN Quartar Installment', 'CNE Subscriber', 'MCE staff (CNE staff)',
                        'BeIN sports CC', 'beIN Bi Installment', 'Corporate Subscriber', 'Temp',
                        'Bein NC', 'Bulk DTH customer', 'beIN Installment Sub', 'Charge Back']
        plans_to_exclude = ['AFCON23ADDNov','EURO24ADDNov','AFC23-ADD-Nov','AFCON23-SA-Nov','AFC23-SA-Nov','EURO24SANov','Temp01']
        dth = df.loc[df['Status']=='Active']
        dth = dth.loc[dth['Customer Type'].isin(dth_sub_types)]
        dth = dth.loc[~dth['Plan'].isin(plans_to_exclude)]
        dth['End Date'] = pd.to_datetime(dth['End Date'])
        dth = dth.sort_values(['Customer Number','End Date'],ascending=False)
        dth = dth.drop_duplicates(subset=['Customer Number'],keep='first')
        dth = dth.loc[:,['Customer Number',
        'Customer Type',
        'Plan',
        'Decoder',
        'Item Description STB',
        'Smart Card',
        'Item Description SC']]
        dth.to_csv(f'{base_dir}/{date} {dth_file_name}.csv',index=False)