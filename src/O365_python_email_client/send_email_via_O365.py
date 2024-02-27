# https://kontext.tech/column/python/795/python-send-email-via-microsoft-graph-api
# https://docs.microsoft.com/en-us/previous-versions/office/office-365-api/api/version-2.0/use-outlook-rest-api
# https://github.com/AzureAD/microsoft-authentication-library-for-python
# https://github.com/microsoftgraph/microsoft-graph-docs/tree/main/api-reference

""" Send email via O365 """

import click
import msal
import base64
import requests
import jwt
import json
import os
from   asn1crypto import pem, x509
import hashlib

def base64UrlDecode(base64Url):
    padding = '=' * (4 - (len(base64Url) % 4))
    return base64.urlsafe_b64decode(base64Url + padding)

def acquire_jwt_token(keyfile, pubfile, passphrase_file, azure_ad_file):
    """
    Acquire jwt_token using MSAL API
    """

    ############
    # TODO Verify that we are communicating with Microsoft
    # https://gist.github.com/dlenski/fc42156c00a615f4aa18a6d19d67e208
    ###

    with open(azure_ad_file, 'r') as fd_az, \
            open(passphrase_file, 'r') as fd_pass, \
            open(keyfile,'r') as fd_priv, \
            open(pubfile,'rb') as fd_pub:
        pub_pem  = fd_pub.read()
        priv_pem = fd_priv.read()
        passphrase = fd_pass.read()
        azure_ad   = json.loads(fd_az.read())
        #{appclientID:appclientID,tenantID:tenantID}
        client_id  = azure_ad['appclientID']
        tenant_id  = azure_ad['tenantID']

    type_name, headers, pub_der = pem.unarmor(pub_pem)
    pub_cert = x509.Certificate.load(pub_der)
    #print (hashlib.sha1(pub_cert.dump()).digest().hex().upper())
    #print(type_name)
    #print(headers)
    #print(pub_cert.signature_algo)
 
    #print(hashlib.sha1(pub_cert.dump()).digest().hex().upper())
    #print('020D3C83E30502A794B847E4F71EFEE57700C29B')
    #print('Verify that the above match')

    my_cred = None

    if 'rsassa_pkcs1v15' == pub_cert.signature_algo :
        my_cred = {
          'private_key'         : priv_pem
          ,'thumbprint'         : hashlib.sha1(pub_cert.dump()).digest().hex().upper()
          ,'public_certificate' : pub_pem.decode('utf-8')
          ,'passphrase'         : passphrase
        }
    elif 'ecdsa' == pub_cert.signature_algo :
        my_cred = {
          'private_key'         : priv_pem
          ,'thumbprint'         : hashlib.sha1(pub_cert.dump()).digest().hex().upper()
          ,'public_certificate' : pub_pem.decode('utf-8')
        }

    authority_url = f"https://login.microsoftonline.com/{tenant_id}"

    app = msal.ConfidentialClientApplication(
       client_id         = client_id
      ,authority         = authority_url
      ,client_credential = my_cred
    )

    # https://docs.microsoft.com/en-us/python/api/msal/msal.application.confidentialclientapplication?view=azure-python
    jwt_token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    return jwt_token

# ---------------------------------------------------------------------------------------------------------------------------------------------
#todo Consider creating a consolidated credentials file
@click.command()
@click.option( "--azure_ad_file",         default=None,                                   help='JSON file {appclientID:appclientID,tenantID:tenantID}')
# ---------------------------------------------------------------------------------------------------------------------------------------------
@click.option( "--passphrase_file",       default=None, hide_input=True,                  help='private key passphrase file')
@click.option( "--keyfile",               default=None,                                   help='pem key filename')
@click.option( "--pubfile",               default=None,                                   help='pem pub filename')
# ---------------------------------------------------------------------------------------------------------------------------------------------
@click.option( "--debug/--no-debug",      default=None,                                   help='debug flag')
# ---------------------------------------------------------------------------------------------------------------------------------------------
@click.option( "--from_user",             default='',                                     help='add - from')
@click.option( "--to",                    default=None, multiple='True',                  help='add - to ')
@click.option( "--cc",                    default=None, multiple='True',                  help='add - cc')
@click.option( "--bcc",                   default=None, multiple='True',                  help='add - bcc')
@click.option( "--replyto",               default=None,                                   help='add - replyto')
@click.option( "--body_file",             default=None,                                   help='add - body filename')
@click.option( "--body", "-b",            default=None,                                   help='add - body')
@click.option( "--subject_line", "-s",    default=None,                                   help='add - subject')
@click.option( "--attachment_file", "-a", default=None, multiple='True',                  help='attachment filename')
@click.option( "--content_type",          default='Text',                                 help='content type')
@click.option( "--importance", "-i",      default="normal",                               help='message importance (high, normal, low)')
# ---------------------------------------------------------------------------------------------------------------------------------------------
def send_email_via_O365(passphrase_file, from_user, to, cc, bcc, replyto, body_file, body,  subject_line, content_type, debug, attachment_file, keyfile, pubfile, azure_ad_file, importance):
    """
    A 'quick and dirty' Exchange Online Email client\n
    """
    # Prerequisites:
    # MS Graph (https://github.com/microsoftgraph),
    # MSAL (https://learn.microsoft.com/en-us/entra/msal/)
    # office-exchange-online (https://learn.microsoft.com/en-us/exchange/exchange-online)
    # REST grammar https://docs.microsoft.com/en-us/graph/api/resources/message?view=graph-rest-1.0

    if keyfile is None or pubfile is None or azure_ad_file is None:
        ctx = click.get_current_context()
        click.echo(ctx.get_help())
        ctx.exit()
        exit(-1)

    jwt_token = acquire_jwt_token(keyfile, pubfile, passphrase_file, azure_ad_file)

    if "access_token" in jwt_token:
        if debug is not None:
            #print('jwt token aquired')
            print(jwt_token)
            print('https://docs.microsoft.com/en-us/azure/active-directory/develop/id-tokens')
            print(jwt.decode(jwt_token["access_token"],options={'verify_signature': False}))

        userId = from_user
        access_token = jwt_token['access_token']
        endpoint = f'https://graph.microsoft.com/v1.0/users/{userId}/SendMail'

        message = dict()

        if body_file is not None:
            if os.path.isfile(body_file):
                with open(body_file, 'r') as f:
                    body = f.read()

        if subject_line is not None:
            message.update( {'subject' : subject_line})

        if body is not None:
            message.update( {'body': { 'contentType' : content_type ,'content': body.replace('\\n', '\n').replace('\\t', '\t') }})

        to_list=[]
        if to is not None:
            for person in to:
                to_list.append({ 'emailAddress' : { 'address': person }})
            message.update({'toRecipients': to_list})

        cc_list=[]
        if cc is not None:
            for person in cc:
                cc_list.append({ 'emailAddress' : { 'address': person }})
            message.update({'ccRecipients': cc_list})

        bcc_list=[]
        if bcc is not None:
            for person in bcc:
                bcc_list.append({ 'emailAddress' : { 'address': person }})
            message.update({'bccRecipients': bcc_list})

        if importance is not None:
            message.update({'importance': importance})

        replyto_list = []
        if replyto is not None:
            replyto_list.append({'emailAddress' : { 'address': replyto }})
            message.update({'replyTo' :  replyto_list})

        # https://developer.microsoft.com/en-us/graph/graph-explorer
        # https://docs.microsoft.com/en-us/graph/api/resources/attachment?view=graph-rest-1.0
        attachment =''
        if attachment_file is not None:
            attachments = []
            for afile in attachment_file:
                if os.path.isfile(afile):
                    with  open(afile, 'rb') as f:
                        attachment_base64 =  base64.b64encode(f.read())
                    attachments.append( {
                        '@odata.type' : '#microsoft.graph.fileAttachment'
                        ,'contentBytes': attachment_base64.decode('utf-8')
                        ,'name'       : os.path.basename(afile)
                        ,'contentType': 'text/plain'
                    })

            message.update({'attachments' :  attachments})
            message.update({'hasAttachments' : True})

            if debug is not None:
              print(attachment)
              print()

        email_msg =  { 'Message' : message
          ,'saveToSentItems' : 'true'
        }

        if debug is not None:
          print(json.dumps(email_msg))
          print()

        r = requests.post(endpoint, headers={'Authorization' : 'Bearer ' + access_token}, json=email_msg)

        if r.ok:
            print('Email sent successfully')
        else:
            print(r.json())

    else:
        print(jwt_token.get("error"))
        print(jwt_token.get("error_description"))
        print(jwt_token.get("correlation_id"))  # You may need this when reporting a bug

if __name__ == "__main__":
    send_email_via_O365()
