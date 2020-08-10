from exchangelib import DELEGATE, IMPERSONATION, Account, Credentials, Configuration
'''
credentials = Credentials(
    username = 'MYWINDOMAIN\\myusername', #or myusername
    password = 'password'
)
'''

# defining credentials
credentials = Credentials(
    username = 'myusername@email.com', #or myusername
    password = 'password'
)

# config for direct connection
config = Configuration(server='outlook.office365.com', credentials=credentials)

# instantiate an account
test_account = Account(
    primary_smtp_address = 'myusername@email.com',
    config = config,
    autodiscover = False,
    access_type = DELEGATE
)

# Print first 100 inbox messages in reverse order
for item in test_account.inbox.all().order_by('-datetime_received')[:100]:
    print(item.subject, item.body, item.attachments)
