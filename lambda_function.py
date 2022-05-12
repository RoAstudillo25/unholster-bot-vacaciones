import urllib
import boto3

from botVacaciones import main

def lambda_handler(myevent, context):

    s3 = boto3.client('s3')
    key = urllib.parse.unquote_plus(myevent['Records'][0]['s3']['object']['key'], encoding='utf-8')
    response = s3.get_object(Bucket='rexmas-correos', Key=key)
    body = response["Body"].read()
    result = main(body)
    
    #TODO implement
    return {
        'statusCode':200,
        'message':result
        }