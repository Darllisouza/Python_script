#BUCKET - ARQUIVO RETORNO

import boto3
import datetime
import os

s3 = boto3.client('s3', aws_access_key_id='', 
                  aws_secret_access_key='')
print('Conexão realizada!')

bucket_name = 'utilities-fs'
folders = ['NEOENERGIA/']

data_do_arquivo = datetime.datetime(2023, 11, 2)

for folder in folders:
    last_folder = None
    continuation_token = None 
    while True:
        if folder != last_folder:    
            response = s3.list_objects_v2(Bucket=bucket_name, Prefix=folder)
            last_folder = folder
        else:
            response = s3.list_objects_v2(Bucket=bucket_name, Prefix=folder, ContinuationToken=continuation_token)
        objects = response.get('Contents', [])
        for obj in objects:
            last_modified = obj['LastModified'].strftime('%Y-%m-%d')
            key = obj['Key']
            filename = os.path.basename(key)
            if datetime.datetime.strptime(last_modified, '%Y-%m-%d') >= data_do_arquivo:
                i = 1
                while os.path.exists(f"{i}_{filename}"):
                    i += 1
                s3.download_file(bucket_name, key, f"{i}_{filename}")
                print(f"Arquivo {filename} baixado com sucesso.")
        if 'NextContinuationToken' in response:
            continuation_token = response['NextContinuationToken']
        else:
            break
print('Download de arquivos concluído.')


