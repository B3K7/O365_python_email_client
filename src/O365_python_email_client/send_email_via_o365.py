# https://kontext.tech/column/python/795/python-send-email-via-microsoft-graph-api
# https://docs.microsoft.com/en-us/previous-versions/office/office-365-api/api/version-2.0/use-outlook-rest-api
# https://github.com/AzureAD/microsoft-authentication-library-for-python
# https://github.com/microsoftgraph/microsoft-graph-docs/tree/main/api-reference

""" Send email via O365 """

import os
import base64
import json
import hashlib
import requests
import jwt
import click
import msal
from   asn1crypto import pem, x509

def acquire_jwt_token(azure_ad_file, debug):
    """
    Acquire jwt_token using MSAL API
    """

    ############
    # Verify that we are communicating with Microsoft
    # https://gist.github.com/dlenski/fc42156c00a615f4aa18a6d19d67e208
    ###

    azure_ad = None
    priv_pem = None
    pub_cert = None
    passphrase = None
    client_id = None
    tenant_id = None
    with open(azure_ad_file, 'r', encoding='utf-8') as fd_az:
        azure_ad   = json.load(fd_az)
        #{appclientID:appclientID,tenantID:tenantID}
        client_id  = azure_ad['appclientID']
        tenant_id  = azure_ad['tenantID']

        if 'pubfile' in azure_ad:
            with open(azure_ad['pubfile'],'rb', encoding='utf-8') as fd_pub:
                pub_pem  = fd_pub.read()
            type_name, headers, pub_der = pem.unarmor(pub_pem)
            pub_cert = x509.Certificate.load(pub_der)

        if 'passphrase' in azure_ad:
            passphrase = azure_ad['passphrase']

        if 'passphrasefile' in azure_ad:
            with open(azure_ad['passphrase_file'], 'r', encoding='utf-8') as fd_pass:
                passphrase = fd_pass.read()


        if 'keyfile' in azure_ad:
            with open(azure_ad['keyfile'],'r', encoding='utf-8') as fd_priv:
                priv_pem = fd_priv.read()


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
    if debug:
        print(authority_url)

    app = msal.ConfidentialClientApplication(
       client_id         = client_id
      ,authority         = authority_url
      ,client_credential = my_cred
    )

    # https://docs.microsoft.com/en-us/python/api/msal/msal.application.confidentialclientapplication?view=azure-python
    jwt_token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    return jwt_token

# ---------------------------------------------------------------------------------------------------------------------------------------------
@click.command()
@click.option( "--azure_ad_file",         default=None,                                   help='JSON file {appclientID:appclientID,tenantID:tenantID}')
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
def send_email_via_o365(from_user=None, to=None, cc=None, bcc=None, \
                        replyto=None, body_file=None, body=None, \
                        subject_line=None, content_type=None, debug=None, \
                        attachment_file=None, azure_ad_file=None,  \
                        importance=None):
    """
    A 'quick and dirty' Exchange Online Email client\n
    """
    # Prerequisites:
    # MS Graph (https://github.com/microsoftgraph),
    # MSAL (https://learn.microsoft.com/en-us/entra/msal/)
    # office-exchange-online (https://learn.microsoft.com/en-us/exchange/exchange-online)
    # REST grammar https://docs.microsoft.com/en-us/graph/api/resources/message?view=graph-rest-1.0

    if azure_ad_file is None:
        ctx = click.get_current_context()
        click.echo(ctx.get_help())
        ctx.exit()
        exit(-1)

    jwt_token = acquire_jwt_token(azure_ad_file, debug)

    if "access_token" in jwt_token:
        if debug is not None:
            #print('jwt token aquired')
            print(jwt_token)
            print('https://docs.microsoft.com/en-us/azure/active-directory/develop/id-tokens')
            print(jwt.decode(jwt_token["access_token"],options={'verify_signature': False}))

        userId = from_user
        access_token = jwt_token['access_token']
        endpoint = f'https://graph.microsoft.com/v1.0/users/{userId}/SendMail'

        message = {}

        if body_file is not None:
            if os.path.isfile(body_file):
                with open(body_file, 'r', encoding='utf-8') as f:
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

        r = requests.post(endpoint, headers={'Authorization' : 'Bearer ' + access_token}, json=email_msg, timeout=3)

        if r.ok:
            print('Email sent successfully')
        else:
            print(r.json())

    else:
        print(jwt_token.get("error"))
        print(jwt_token.get("error_description"))
        print(jwt_token.get("correlation_id"))  # You may need this when reporting a bug

if __name__ == "__main__":
    send_email_via_o365()
