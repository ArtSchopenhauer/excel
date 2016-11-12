import openpyxl
from collections import OrderedDict
from simple_salesforce import Salesforce

sf = Salesforce(
    username='integration@levelsolar.com',
    password='HrNt7DrqaEfZmBqJRan9dKFzmQFp',
    security_token='yWlJG8lAKCq1pTBkbBMSVcKg')

metrics = sf.query_all("SELECT Date__c, Ambassador__c, Office__c, AmbShifts__c, Shift_Length__c, Doors__c, Appointments__c FROM Metrics__c WHERE Ambassador__c != null AND Office__c != null AND Date__c = (LAST_N_WEEKS:4 OR THIS_WEEK)")["records"]

