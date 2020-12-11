import json
import uuid
import xlrd
import os
import xlwt
import xlsxwriter
import openpyxl
import os.path
import datetime
from datetime import datetime
import requests

class licenses():
    def __init__(self,id):
        self.id=id
        # licenses.licencePrint(self, ermName, vendor[0], vendor[1],ermDescription, ERMlicType, ERMlicstatus, ErmstartDate, ermendDate,docs)

    def licencePrint(self, ermName, vendor,ermDescription, ERMlicType, ERMlicstatus, ErmstartDate, ermendDate,docs,aliases,consortium,ermopenEnded,cp,path):
        Ordarchivo=open(path, 'a')
        if (consortium):
            license = {
                #"id": "1ece2ebf-7999-4720-8399-d3c04022647c",
                #"dateCreated": "2020-11-29T15:33:02Z",
                "links": [],
                "description": ermDescription,
                "customProperties": cp,
                "contacts": [],
                "tags": [],
                "lastUpdated": "2020-11-29T15:33:02Z",
                "docs":docs,
                "name": ermName,
                "status": ERMlicstatus,
                "supplementaryDocs": [],
                "startDate": ErmstartDate,
                "endDate": ermendDate,
                "_links": {"linkedResources": {"href": "/licenses/licenseLinks?filter=owner.id%"}},
                "openEnded": ermopenEnded,
                "amendments": [],
                "orgs": [{"id": "","org": {"id": "","orgsUuid": vendor[0],"name": vendor[1]},"role": {"id": "","value": "licensor","label": "Licensor"}},{ "id": "","org": {"id": "","orgsUuid": consortium[0],"name": consortium[1]},"role": {"id": "","value": "consortium","label": "consortium"}}],
                "type": ERMlicType,
                "alternateNames": aliases
                }
        else:
            license = {
                #"id": "1ece2ebf-7999-4720-8399-d3c04022647c",
                #"dateCreated": "2020-11-29T15:33:02Z",
                "links": [],
                "description": ermDescription,
                "customProperties": cp,
                "contacts": [],                
                "tags": [],
                "lastUpdated": "2020-11-29T15:33:02Z",
                "docs":docs,
                "name": ermName,
                "status": ERMlicstatus,
                "supplementaryDocs": [],
                "startDate": ErmstartDate,
                "endDate": ermendDate,
                "_links": {"linkedResources": {"href": "/licenses/licenseLinks?filter=owner.id%"}},
                "openEnded": ermopenEnded,
                "amendments": [],
                "orgs": [{"id": "","org": {"id": "","orgsUuid": vendor[0],"name": vendor[1]},"role": {"id": "","value": "licensor","label": "Licensor"}}],
                "type": ERMlicType,
                "alternateNames": aliases
                }

        #json_ord = json.dumps(order,indent=2)
        json_ord = json.dumps(license)
        print('Datos en formato JSON', json_ord)
        Ordarchivo.write(json_ord+"\n")

def print_notes(title,linkId,cont):
    Ordarchivo=open("widener\licenses\widener_notes.json", 'a')
    tn="eb1e0b21-69a6-4b99-91db-4fcb10a7fca2"
    notes ={
           "typeId": tn,
           "type": "General note",
           "domain": "licenses",
           "title": title,
           "content": "<p>"+cont+"</p>",
           "links": [{"id": linkId[0],"type": "license"}]
           }
    json_ord = json.dumps(notes,indent=2)
    json_notes = json.dumps(notes)
    print('Datos en formato JSON', json_notes)
    Ordarchivo.write(json_notes+"\n")

def get_licId_Macewan(orgname):
        dic={}
        #pathPattern="/organizations-storage/organizations" #?limit=9999&query=code="
        #https://okapi-macewan.folio.ebsco.com/licenses/licenses?stats=true&term=Teatro Español del Siglo de Oro&match=name
        pathPattern="/licenses/licenses" #?limit=9999&query=code="
        okapi_url="https://okapi-macewan.folio.ebsco.com"
        okapi_token="eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJhZG1pbiIsInVzZXJfaWQiOiI4MjEzODdhZS1hNzkxLTQ5NTgtYTg3ZS1jYTFmMDE2NzA2YmUiLCJpYXQiOjE2MDY2NTg2NzcsInRlbmFudCI6ImZzMDAwMDEwMzcifQ.YFlctF1WxhO_f-Sc0_KIi_UCD5cngon5wZ6rCpgPLEA"
        okapi_tenant="fs00001037"
        okapi_headers = {"x-okapi-token": okapi_token,"x-okapi-tenant": okapi_tenant,"content-type": "application/json"}
        length="1"
        start="1"
        element="organizations"
        query=f"?stats=true&term="
        #/organizations-storage/organizations?query=code==UMPROQ
        paging_q = f"{query}"+orgname+"&match=name"
        path = pathPattern+paging_q
        #data=json.dumps(payload)
        url = okapi_url + path
        req = requests.get(url, headers=okapi_headers)
        idorg=[]
        if req.status_code != 201:
            json_str = json.loads(req.text)
            total_recs = int(json_str["totalRecords"])
            if (total_recs!=0):
                #print('Datos en formato JSON',json.dumps(json_str))
                rec=json_str["results"]
                #print(json_str)
                l=rec[0]
                if 'id' in l:
                    idorg.append(l['id'])
                    #idorg.append(l['name'])
        if len(idorg)==0:
            return "00000-000000-000000-00000"
        else:
            return idorg
   

def floatHourToTime(fh):
    h, r = divmod(fh, 1)
    m, r = divmod(r*60, 1)
    return (
        int(h),
        int(m),
        int(r*60),
    )

def getorgid_Macewan(orgname):
        dic={}
        #pathPattern="/organizations-storage/organizations" #?limit=9999&query=code="
        pathPattern="/organizations/organizations" #?limit=9999&query=code="
        okapi_url="https://okapi-macewan.folio.ebsco.com"
        okapi_token="eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJhZG1pbiIsInVzZXJfaWQiOiI4MjEzODdhZS1hNzkxLTQ5NTgtYTg3ZS1jYTFmMDE2NzA2YmUiLCJpYXQiOjE2MDYxNDU2NzgsInRlbmFudCI6ImZzMDAwMDEwMzcifQ.vvHnTq48ERXtWKmSl4vt7Tlm7p11Gp7ge6XKc3EBJCA"
        okapi_tenant="fs00001037"
        okapi_headers = {"x-okapi-token": okapi_token,"x-okapi-tenant": okapi_tenant,"content-type": "application/json"}
        length="1"
        start="1"
        element="organizations"
        query=f"query=name=="
        #/organizations-storage/organizations?query=code==UMPROQ
        paging_q = f"?{query}"+'"'+f"{orgname}"+'"'
        path = pathPattern+paging_q
        #data=json.dumps(payload)
        url = okapi_url + path
        req = requests.get(url, headers=okapi_headers)
        idorg=[]
        if req.status_code != 201:
            json_str = json.loads(req.text)
            total_recs = int(json_str["totalRecords"])
            if (total_recs!=0):
                rec=json_str[element]
                #print(rec)
                l=rec[0]
                if 'id' in l:
                    idorg.append(l['id'])
                    idorg.append(l['name'])
            if len(idorg)==0:
                idorg.append("0b3ffc1e-1fa5-4d40-8c93-e592aa94ab57")
                idorg.append("EBSCO")
        return idorg

def getorgid_Widener(orgname):
        dic={}
        #pathPattern="/organizations-storage/organizations" #?limit=9999&query=code="
        pathPattern="/organizations/organizations" #?limit=9999&query=code="
        okapi_url="https://okapi-widener.folio.ebsco.com"
        okapi_token="eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJhZG1pbiIsInVzZXJfaWQiOiI2NjU3ZTFlOS04M2E3LTQ3ZDEtOTEyOS03ZDY2ZDY1NzYyMWIiLCJpYXQiOjE2MDc2MTc4OTQsInRlbmFudCI6ImZzMDAwMDEwMzgifQ.rIIuUkPchhf7wLxASOel37OngoM-HasQj6SyKKFjBR4"
        okapi_tenant="fs00001038"
        okapi_headers = {"x-okapi-token": okapi_token,"x-okapi-tenant": okapi_tenant,"content-type": "application/json"}
        length="1"
        start="1"
        element="organizations"
        query=f"query=name=="
        #/organizations-storage/organizations?query=code==UMPROQ
        paging_q = f"?{query}"+'"'+f"{orgname}"+'"'
        path = pathPattern+paging_q
        #data=json.dumps(payload)
        url = okapi_url + path
        req = requests.get(url, headers=okapi_headers)
        idorg=[]
        if req.status_code != 201:
            json_str = json.loads(req.text)
            total_recs = int(json_str["totalRecords"])
            if (total_recs!=0):
                rec=json_str[element]
                #print(rec)
                l=rec[0]
                if 'id' in l:
                    idorg.append(l['id'])
                    idorg.append(l['name'])
            if len(idorg)==0:
                idorg.append("0b3ffc1e-1fa5-4d40-8c93-e592aa94ab57")
                idorg.append("EBSCO")
        return idorg
def getorgid_Liverpool(orgname):
        dic={}
        #pathPattern="/organizations-storage/organizations" #?limit=9999&query=code="
        pathPattern="/organizations/organizations" #?limit=9999&query=code="
        okapi_url="https://okapi-liverpool.folio.ebsco.com"
        okapi_token="eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJhZG1pbiIsInVzZXJfaWQiOiI2NGEyZWY0Yy04YjBkLTRlMjYtYmU3Yy1jOWNkNmM4MTYwYmMiLCJpYXQiOjE2MDY2NzY3OTMsInRlbmFudCI6ImZzMDAwMDEwNDUifQ.WG8OXJMcq4-GUTzaLkA4CjKAkZcl98GG2qQ3vD-yCr0"
        okapi_tenant="fs00001045"
        okapi_headers = {"x-okapi-token": okapi_token,"x-okapi-tenant": okapi_tenant,"content-type": "application/json"}
        length="1"
        start="1"
        element="organizations"
        query=f"query=name=="
        #/organizations-storage/organizations?query=code==UMPROQ
        paging_q = f"?{query}"+'"'+f"{orgname}"+'"'
        path = pathPattern+paging_q
        #data=json.dumps(payload)
        url = okapi_url + path
        req = requests.get(url, headers=okapi_headers)
        idorg=[]
        if req.status_code != 201:
            json_str = json.loads(req.text)
            total_recs = int(json_str["totalRecords"])
            if (total_recs!=0):
                rec=json_str[element]
                #print(rec)
                l=rec[0]
                if 'id' in l:
                    idorg.append(l['id'])
                    idorg.append(l['name'])
            if len(idorg)==0:
                idorg.append("5dc41889-f566-417b-a5e7-36e686e04ee1")
                idorg.append("undefined")
        return idorg

def get_licId_Liverpool(orgname):
        dic={}
        #pathPattern="/organizations-storage/organizations" #?limit=9999&query=code="
        #https://okapi-macewan.folio.ebsco.com/licenses/licenses?stats=true&term=Teatro Español del Siglo de Oro&match=name
        pathPattern="/licenses/licenses" #?limit=9999&query=code="
        okapi_url="https://okapi-liverpool.folio.ebsco.com"
        okapi_token="eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJhZG1pbiIsInVzZXJfaWQiOiI2NGEyZWY0Yy04YjBkLTRlMjYtYmU3Yy1jOWNkNmM4MTYwYmMiLCJpYXQiOjE2MDY2NzY3OTMsInRlbmFudCI6ImZzMDAwMDEwNDUifQ.WG8OXJMcq4-GUTzaLkA4CjKAkZcl98GG2qQ3vD-yCr0"
        okapi_tenant="fs00001045"
        okapi_headers = {"x-okapi-token": okapi_token,"x-okapi-tenant": okapi_tenant,"content-type": "application/json"}
        length="1"
        start="1"
        element="organizations"
        query=f"?stats=true&term="
        #/organizations-storage/organizations?query=code==UMPROQ
        paging_q = f"{query}"+orgname+"&match=name"
        path = pathPattern+paging_q
        #data=json.dumps(payload)
        url = okapi_url + path
        req = requests.get(url, headers=okapi_headers)
        idorg=[]
        if req.status_code != 201:
            json_str = json.loads(req.text)
            total_recs = int(json_str["totalRecords"])
            if (total_recs!=0):
                #print('Datos en formato JSON',json.dumps(json_str))
                rec=json_str["results"]
                #print(json_str)
                l=rec[0]
                if 'id' in l:
                    idorg.append(l['id'])
                    #idorg.append(l['name'])
        if len(idorg)==0:
            return "00000-000000-000000-00000"
        else:
            return idorg


def get_licId_Wineder(orgname):
        dic={}
        #pathPattern="/organizations-storage/organizations" #?limit=9999&query=code="
        #https://okapi/licenses/licenses?stats=true&term=Teatro Español del Siglo de Oro&match=name
        pathPattern="/licenses/licenses" #?limit=9999&query=code="
        okapi_url="https://okapi-widener.folio.ebsco.com"
        okapi_token="eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJhZG1pbiIsInVzZXJfaWQiOiI2NjU3ZTFlOS04M2E3LTQ3ZDEtOTEyOS03ZDY2ZDY1NzYyMWIiLCJpYXQiOjE2MDc2MTc4OTQsInRlbmFudCI6ImZzMDAwMDEwMzgifQ.rIIuUkPchhf7wLxASOel37OngoM-HasQj6SyKKFjBR4"
        okapi_tenant="fs00001038"
        okapi_headers = {"x-okapi-token": okapi_token,"x-okapi-tenant": okapi_tenant,"content-type": "application/json"}
        length="1"
        start="1"
        element="organizations"
        query=f"?stats=true&term="
        #/organizations-storage/organizations?query=code==UMPROQ
        paging_q = f"{query}"+orgname+"&match=name"
        path = pathPattern+paging_q
        #data=json.dumps(payload)
        url = okapi_url + path
        req = requests.get(url, headers=okapi_headers)
        idorg=[]
        if req.status_code != 201:
            json_str = json.loads(req.text)
            total_recs = int(json_str["totalRecords"])
            if (total_recs!=0):
                #print('Datos en formato JSON',json.dumps(json_str))
                rec=json_str["results"]
                #print(json_str)
                l=rec[0]
                if 'id' in l:
                    idorg.append(l['id'])
                    #idorg.append(l['name'])
        if len(idorg)==0:
            return "00000-000000-000000-00000"
        else:
            return idorg

def licType(a,cust):
    lt={}
    if cust=="M":
        if (a=="Signed License Agreement"):
            lt={"id": "2c918085744529450175654f44830000","value": "signed_license","label": "Signed License Agreement"}
        elif (a=="Passive Assent License Agreement"):
            lt={"id": "2c91808574452945017565504adf0002","value": "passive_assent_license_agreement","label": "Passive Assent License Agreement"}
        elif(a=="Site Terms And Conditions"):
            lt={"value": "site_terms_and_conditions","label": "site_terms_and_conditions"}
        else:
            lt={"id": "","value": "local","label": "Local"}
    ##Liverpool
    elif cust=="L":
        if (a=="Signed License Agreement"):
            lt={"id": "2c91808d7426454d017615dc546e0004","value": "signed_license","label": "Signed License Agreement"}
        elif (a=="Site Terms And Conditions"):
            lt={"id": "2c91808d7426454d01761f9410c50022","value": "unspecified","label": "unspecified"}
        else:
            lt={"id": "","value": "local","label": "Local"}

    elif cust=="W":
        if (a=="Signed License Agreement"):
            lt={"id": "2c9180857445294501764d907904005e","value": "signed_license_agreement","label": "Signed License Agreement"}
        elif (a=="Passive Assent License Agreement"):
            lt={"id": "2c91808b7445501601764d90390f0033","value": "passive_assent_license_agreement","label": "Passive Assent License Agreement"}
        elif(a=="Site Terms And Conditions"):
            lt={"id": "2c9180857445294501764d909457005f","value": "site_terms_and_conditions","label": "site_terms_and_conditions"}
        else:
            lt={"id": "","value": "local","label": "Local"}

    return lt

def date_stamp(ilsdate):
    dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(ilsdate) - 2)
    hour, minute, second = floatHourToTime(ilsdate % 1)
    dt = str(dt.replace(hour=hour, minute=minute,second=second))+".000+0000" #Approbal by
    dia=dt[8:10]
    mes=dt[5:7]
    ano=dt[0:4]
    dt=ano+"-"+mes+"-"+dia
    return dt


                #End Date
###################################################################################################
#
#
###################################################################################################
def readSpreadsheet(spreadsheet,path,org,orgL):
    wb = xlrd.open_workbook(spreadsheet)
    fileN=org
    worksheet = wb.sheet_by_name("all")
    print("no rows: ", worksheet.nrows)
    print("no colomns: ", worksheet.ncols)
    f = open("licenses_error.txt", "a")
    #Ordarchivo=open("macewan\licenses\macewan_licenses.json", 'a')
    Ordarchivo=open(path, 'a')
    #read orders
    count=0
    oldlicense=""
    diccionario={}
    for p in range(worksheet.nrows):
        if (p!=0):
                if p==68:
                    stop=1

                database=worksheet.cell_value(p,1)
            #if oldlicense!=database:
                orgname=worksheet.cell_value(p,1)
                #idLic=get_licId(orgname)
                year=[]
                print("############### Record No"+str(p)+"##################\n")
                oldlicense=worksheet.cell_value(p,1)
                #License Year
                ermLicenceYear=str(worksheet.cell_value(p,0))
                ermLicenceYear=ermLicenceYear.replace(".0","")
                #License Name
                ermName=str(worksheet.cell_value(p,1).strip())
                #Organization Name
                vendor=[]
                if fileN=="Macewan_licenses":
                    if (worksheet.cell_value(p,2)=="Proquest Info Learning Co"):
                        vendor.append("19fdabf3-b73a-4da1-bb5b-25e4aa737ab7")
                        vendor.append("Proquest Info Learning Co")
                    else:
                        vendor=getorgid_Macewan(worksheet.cell_value(p,2).strip())
                if fileN=="Liverpool_licenses":
                    if worksheet.cell_value(p,2):
                        vendor=getorgid_Liverpool(worksheet.cell_value(p,2).strip())
                    else:
                        idorg.append("5dc41889-f566-417b-a5e7-36e686e04ee1")
                        idorg.append("undefined")
                    
                if fileN=="Widener_licenses":
                    vendor=getorgid_Widener(worksheet.cell_value(p,2).strip())

                consortia=""
                if worksheet.cell_value(p,4):
                    consortia=getorgid_Macewan(worksheet.cell_value(p,4).strip())
                #Aliases
                erm_aliases=worksheet.cell_value(p,4).strip()
                #License Type- Local by default
                ERMlicType=licType(worksheet.cell_value(p,23).strip(),orgL)
                
                #License Status by default active
                if fileN=="Widener_licenses":
                    ERMlicstatus={"id": "2c91808f725c72b30172718dc70d0045","value": "active","label": "Active"}    
                elif fileN=="Macewan_licenses":
                    ERMlicstatus={"id": "2c91808f725c72b30172622a26e20006","value": "active","label": "Active"}    
                elif fileN=="Liverpoool_licenses":
                    ERMlicstatus={"id": "2c9180837422170501742226631d000b","value": "active","label": "Active"}    
                #Start Date
                ermStartDate=""
                if worksheet.cell_value(p,20):
                    ermStartDate=date_stamp(worksheet.cell_value(p,20))
                else:
                    if (org=="Macewan_licenses"):
                        ermLicenceYear=int(ermLicenceYear)-1
                        ermStartDate=str(ermLicenceYear)+"-07-01"
                #End Date
                ermendDate=""
                ermopenEnded= False
                if worksheet.cell_value(p,17):
                    ermendDate=date_stamp(worksheet.cell_value(p,17))
                else:
                    ermendDate=""
                    ermopenEnded= True
                #Description
                ermDescriptionA=""
                if worksheet.cell_value(p,10):
                    ermDescriptionA=str(worksheet.cell_value(p,10))
                docs=[]
                if worksheet.cell_value(p,27):
                    url=worksheet.cell_value(p,27).strip()
                    if worksheet.cell_value(p,28):
                        namedoc=worksheet.cell_value(p,28)
                        docs.append({"id": "","dateCreated": "2020-11-23T00:00:00Z","lastUpdated": "2020-11-23T00:00:00Z","url": url,"name": namedoc})
                if worksheet.cell_value(p,29):
                    url=worksheet.cell_value(p,29).strip()
                    if worksheet.cell_value(p,30):
                        namedoc=worksheet.cell_value(p,30)
                        docs.append({"id": "","dateCreated": "2020-11-23T00:00:00Z","lastUpdated": "2020-11-23T00:00:00Z","url": url,"name": namedoc})
                if worksheet.cell_value(p,31):
                    url=worksheet.cell_value(p,31).strip()
                    if worksheet.cell_value(p,32):
                        namedoc=worksheet.cell_value(p,32)
                        docs.append({"id": "","dateCreated": "2020-11-23T00:00:00Z","lastUpdated": "2020-11-23T00:00:00Z","url": url,"name": namedoc})
                if worksheet.cell_value(p,33):
                    url=worksheet.cell_value(p,33).strip()
                    if worksheet.cell_value(p,34):
                        namedoc=worksheet.cell_value(p,34)
                        docs.append({"id": "","dateCreated": "2020-11-23T00:00:00Z","lastUpdated": "2020-11-23T00:00:00Z","url": url,"name": namedoc})
                uuidpol=str(uuid.uuid4())
            ##################ermName, ermOrgid,ermOrgName
                aliases=""
                #databaseNext=worksheet.cell_value(p+1,1).strip()
            #if database!=databaseNext:
######################################################
#
# LICENCE TERMS
#
######################################################
                if orgL=="M":
                    cp={}
                    if worksheet.cell_value(p,5):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,6))
                            label=str(worksheet.cell_value(p,5))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["AllRightsReserved"]= {"note":note,"internal": True,"value": {"value": value,"label": label}} 
                        
                    if worksheet.cell_value(p,6):
                            note=""
                            value=str(worksheet.cell_value(p,6))
                            cp["AuthorizedUsers"]={"note": note, "internal": True,"value": value,"type": {"id": "2c91808574452945017565ac3f550004","name": "AuthorizedUsers","primary": True,"defaultInternal": True,"label": "Authorized User Definition","description": "Defines what constitutes an authorized user for the licensed resource.","weight": 0,"type": "com.k_int.web.toolkit.custprops.types.CustomPropertyText"}}
                        
                    if worksheet.cell_value(p,7):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,8))
                            label=str(worksheet.cell_value(p,7))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["AgreementConfidentiality"]= {"note":note,"internal": True,"value": {"value": value,"label": label}} 
                        

                    if worksheet.cell_value(p,9):
                            label=""
                            value=""
                            note=""
                            label=str(worksheet.cell_value(p,9))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["UserConfidentiality"]= {"internal": True,"value": {"value": value,"label": label}} 
                       
  
                    if worksheet.cell_value(p,11):
                            note=""
                            note=str(worksheet.cell_value(p,11))
                            cp["GovJurisdiction"]= {"internal": True,"value": note,"type": {"name": "GovJurisdiction","primary": True,"defaultInternal": True,"label": "Governing Jurisdiction","description": "Details the governing jurisdiction of the license agreement","weight": 0,"type": "com.k_int.web.toolkit.custprops.types.CustomPropertyText"}}
                        
                        
                    if worksheet.cell_value(p,12):
                            note=""
                            note=str(worksheet.cell_value(p,12))
                            cp["GovLaw"]= {"internal": True,"value": note,"type": {"name": "GovLaw","primary": True,"defaultInternal": True,"label": "Governing Law","description": "Details the governing law of the license agreement","weight": 0,"type": "com.k_int.web.toolkit.custprops.types.CustomPropertyText"}}
                       
                            #note=""
                    if worksheet.cell_value(p,13):
                            note=""
                            note=str(worksheet.cell_value(p,13))
                            cp["IndemLicensee"]={"note": note,"internal": True,"value": "Indemnification by Licensee","type": {"id": "2c91808b74455016017565beec470004","name": "IndemLicensee","primary": True,"defaultInternal": True,"label": "Indemnification by Licensee","description": "Indemnification by Licensee","weight": 0,"type": "com.k_int.web.toolkit.custprops.types.CustomPropertyText"}}
                        
                    if worksheet.cell_value(p,14):
                            note=str(worksheet.cell_value(p,14))
                            cp["IndemLicensor"] ={"note": note,"internal": True,"value": "Indemnification by Licensor","type": {"id": "2c91808b74455016017565bf87920005","name": "IndemLicensor","primary": True,"defaultInternal": True,"label": "Indemnification by Licensor","description": "Indemnification by Licensor","weight": 0,"type": "com.k_int.web.toolkit.custprops.types.CustomPropertyText"}}
                       
                            note=""
                    if worksheet.cell_value(p,15):
                            note=str(worksheet.cell_value(p,15))
                    if worksheet.cell_value(p,24):
                            label=""
                            value=""
                            note=""
                            label=str(worksheet.cell_value(p,24))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["SERU"]= {"internal": True,"value": {"value": value,"label": label}}
                       
                        
                    if worksheet.cell_value(p,35):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,36))
                            label=str(worksheet.cell_value(p,35))
                            value=str(worksheet.cell_value(p,36))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["AlumniAccess"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}
  
                    if worksheet.cell_value(p,37):
                            note=""
                            value=str(worksheet.cell_value(p,37))
                            cp["CopyrightLaw"]={"id": 529,"note": note,"internal": True,"value": value,"type": {"id": "2c91808574452945017565e389c2000a","name": "CopyrightLaw","primary": True,"defaultInternal": True,"label": "Applicable Copyright Law","description": "Specifies the copyright law applicable to the licensed resource","weight": 0,"type": "com.k_int.web.toolkit.custprops.types.CustomPropertyText"}}
 
                        
                    if worksheet.cell_value(p,38):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,39))
                            label=str(worksheet.cell_value(p,38))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["ArchivingAllowed"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}
 
                    if worksheet.cell_value(p,41):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,42))
                            label=str(worksheet.cell_value(p,41))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["ArticleReach"]={ "note":note,"internal": True,"value": {"value": value,"label": label}} 

                    if worksheet.cell_value(p,43):
                            if fileN=="Macewan_licenses":
                                note=str(worksheet.cell_value(p,44))
                                value=int(worksheet.cell_value(p,43))
                            elif fileN=="Liverpool_licenses":
                                note=str(worksheet.cell_value(p,43))
                                value=0
                            cp["ConcurrentUsers"]={"note": note,"internal": True,"value": value,"type": {"id": "2c91808b74455016017566e8792f001f","name": "ConcurrentUsers","primary": True,"defaultInternal": True,"label": "Concurrent Users","description": "Specifies the number of allowed concurrent users","weight": 0,"type": "com.k_int.web.toolkit.custprops.types.CustomPropertyInteger"}}
                        
                    if worksheet.cell_value(p,45):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,46))
                            label=str(worksheet.cell_value(p,45))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["CopyDigital"]={ "note":note,"internal": True,"value": {"value": value,"label": label}} 

                    if worksheet.cell_value(p,47):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,48))
                            label=str(worksheet.cell_value(p,47))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["CopyPrint"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}    
 
                    if worksheet.cell_value(p,49):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,50))
                            label=str(worksheet.cell_value(p,49))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["CoursePackElectronic"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}    

                    if worksheet.cell_value(p,51):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,52))
                            label=str(worksheet.cell_value(p,51))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["CoursePackPrint"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}

                    if worksheet.cell_value(p,53):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,54))
                            label=str(worksheet.cell_value(p,53))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["DistanceEducation"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}

                    if worksheet.cell_value(p,55):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,56))
                            label=str(worksheet.cell_value(p,55))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["FairDealingClause"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}
 
                    if worksheet.cell_value(p,57):
                            label=""
                            value=str(worksheet.cell_value(p,57))
                            note=""
                            note=str(worksheet.cell_value(p,58))
                            label=str(worksheet.cell_value(p,57))
                            #if label=="":
                            #    value="unspecified"
                            #    label="unspecified"
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["FairUseClause"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}
 
                    if worksheet.cell_value(p,59):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,60))
                            label=str(worksheet.cell_value(p,59))
                            value=str(worksheet.cell_value(p,59))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["ILLElectronicSecure"]={ "note":note,"internal": True,"value": {"value": value,"label": label}} 

                        
                    if worksheet.cell_value(p,61):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,62))
                            label=str(worksheet.cell_value(p,61))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["ILLPrint"]={ "note":note,"internal": True,"value": {"value": value,"label": label}} 
  
                    if worksheet.cell_value(p,63):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,64))
                            label=str(worksheet.cell_value(p,63))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["ILLElectronicSecure"]={ "note":note,"internal": True,"value": {"value": value,"label": label}} 
  
                    if worksheet.cell_value(p,65):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,66))
                            label=str(worksheet.cell_value(p,65))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["LinkElectronic"]={ "note":note,"internal": True,"value": {"value": value,"label": label}} 

                    if worksheet.cell_value(p,67):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,68))
                            label=str(worksheet.cell_value(p,67))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["LMS"]={ "note":note,"internal": True,"value": {"value": value,"label": label}} 
 
                    if worksheet.cell_value(p,69):
                            note=""
                            value=str(worksheet.cell_value(p,69))
                            cp["OtherRestrictions"]={"note": note,"internal": True,"value": value,"type": {"name": "OtherRestrictions","primary": True,"defaultInternal": True,"label": "Other Restrictions","description": "A blanket term to capture restrictions on a licensed resource not covered by established terms","weight": 0,"type": "com.k_int.web.toolkit.custprops.types.CustomPropertyText"}}
                            #OthersRestrictions
                    if worksheet.cell_value(p,70):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,71))
                            label=str(worksheet.cell_value(p,70))
                            value=str(worksheet.cell_value(p,70))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["PerpetualAccess"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}

                       
                    if worksheet.cell_value(p,72):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,73))
                            label=str(worksheet.cell_value(p,72))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["PersonsPerceptualDisabilities"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}
 
                       
                    if worksheet.cell_value(p,74):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,75))
                            label=str(worksheet.cell_value(p,74))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["ReservesElectronic"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}
 
                      
                    if worksheet.cell_value(p,76):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,77))
                            label=str(worksheet.cell_value(p,76))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["ReservesPrint"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}

                       
                    if worksheet.cell_value(p,78):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,79))
                            label=str(worksheet.cell_value(p,78))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["DepositRights"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}
 
                       
                    if worksheet.cell_value(p,80):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,81))
                            label=str(worksheet.cell_value(p,80))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["ScholarlySharing"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}
 
                    if worksheet.cell_value(p,82):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,83))
                            label=str(worksheet.cell_value(p,82))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["TextDataMining"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}
  
                       
                    if worksheet.cell_value(p,84):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,85))
                            label=str(worksheet.cell_value(p,84))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["Walkins"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}
################LIVERPOOL###############
#
###############################
                elif orgL=="L":
                    cp={}
                    #if worksheet.cell_value(p,5):
                    #        label=""
                    #        value=""
                    #        note=""
                    #        note=str(worksheet.cell_value(p,6))
                    #        label=str(worksheet.cell_value(p,5))
                    #        value=label.replace(" ","_")
                    #        value=value.lower()
                    #        cp["AllRightsReserved"]= {"note":note,"internal": True,"value": {"value": value,"label": label}} 
                        
                    if worksheet.cell_value(p,6):
                            note=""
                            note=str(worksheet.cell_value(p,6))
                            cp["AuthorizedUsers"]={"note": note, "internal": True,"value": "Authorized User Definition","type": {"id": "2c91808574452945017565ac3f550004","name": "AuthorizedUsers","primary": True,"defaultInternal": True,"label": "Authorized User Definition","description": "Defines what constitutes an authorized user for the licensed resource.","weight": 0,"type": "com.k_int.web.toolkit.custprops.types.CustomPropertyText"}}
                        
                    #if worksheet.cell_value(p,7):
                    #        label=""
                    #        value=""
                    #        note=""
                    #        note=str(worksheet.cell_value(p,8))
                    #        label=str(worksheet.cell_value(p,7))
                    #        value=label.replace(" ","_")
                    #        value=value.lower()
                    #        cp["AgreementConfidentiality"]= {"note":note,"internal": True,"value": {"value": value,"label": label}} 
                        

                    #if worksheet.cell_value(p,9):
                    #        label=""
                    #        value=""
                    #        note=""
                    #        label=str(worksheet.cell_value(p,9))
                    #        value=label.replace(" ","_")
                    #        value=value.lower()
                    #        cp["UserConfidentiality"]= {"internal": True,"value": {"value": value,"label": label}} 
                       
                    
                    #if worksheet.cell_value(p,11):
                    #        note=""
                    #        note=str(worksheet.cell_value(p,11))
                    #        cp["GovJurisdiction"]= {"internal": True,"value": note,"type": {"name": "GovJurisdiction","primary": True,"defaultInternal": True,"label": "Governing Jurisdiction","description": "Details the governing jurisdiction of the license agreement","weight": 0,"type": "com.k_int.web.toolkit.custprops.types.CustomPropertyText"}}
                        
                        
                    #if worksheet.cell_value(p,12):
                    #        note=""
                    #        note=str(worksheet.cell_value(p,12))
                    #        cp["GovLaw"]= {"internal": True,"value": note,"type": {"name": "GovLaw","primary": True,"defaultInternal": True,"label": "Governing Law","description": "Details the governing law of the license agreement","weight": 0,"type": "com.k_int.web.toolkit.custprops.types.CustomPropertyText"}}
                       
                            #note=""
                    #if worksheet.cell_value(p,13):
                    #        note=""
                    #        note=str(worksheet.cell_value(p,13))
                    #        cp["IndemLicensee"]={"note": note,"internal": True,"value": "Indemnification by Licensee","type": {"id": "2c91808b74455016017565beec470004","name": "IndemLicensee","primary": True,"defaultInternal": True,"label": "Indemnification by Licensee","description": "Indemnification by Licensee","weight": 0,"type": "com.k_int.web.toolkit.custprops.types.CustomPropertyText"}}
                        
                    #if worksheet.cell_value(p,14):
                    #        note=str(worksheet.cell_value(p,14))
                    #        cp["IndemLicensor"] ={"note": note,"internal": True,"value": "Indemnification by Licensor","type": {"id": "2c91808b74455016017565bf87920005","name": "IndemLicensor","primary": True,"defaultInternal": True,"label": "Indemnification by Licensor","description": "Indemnification by Licensor","weight": 0,"type": "com.k_int.web.toolkit.custprops.types.CustomPropertyText"}}
                       
                            note=""
                    if worksheet.cell_value(p,15):
                            note=str(worksheet.cell_value(p,15))
                    #if worksheet.cell_value(p,24):
                    #        label=""
                    #        value=""
                    #        note=""
                    #        label=str(worksheet.cell_value(p,24))
                    #        value=label.replace(" ","_")
                    #        value=value.lower()
                    #        cp["SERU"]= {"internal": True,"value": {"value": value,"label": label}}
                       
                        
                    if worksheet.cell_value(p,35):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,36))
                            label=str(worksheet.cell_value(p,35))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["AlumniAccess"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}
  
                    #if worksheet.cell_value(p,37):
                    #        note=""
                    #        note=str(worksheet.cell_value(p,37))
                    #        cp["CopyrightLaw"]={"id": 529,"note": note,"internal": True,"value": "Applicable Copyright Law","type": {"id": "2c91808574452945017565e389c2000a","name": "CopyrightLaw","primary": True,"defaultInternal": True,"label": "Applicable Copyright Law","description": "Specifies the copyright law applicable to the licensed resource","weight": 0,"type": "com.k_int.web.toolkit.custprops.types.CustomPropertyText"}}
 
                        
                    #if worksheet.cell_value(p,38):
                    #        label=""
                    #        value=""
                    #        note=""
                    #        note=str(worksheet.cell_value(p,39))
                    #        label=str(worksheet.cell_value(p,38))
                    #        value=label.replace(" ","_")
                    #        value=value.lower()
                    #        cp["ArchivingAllowed"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}
 
                    #if worksheet.cell_value(p,41):
                    #        label=""
                    #        value=""
                    #        note=""
                    #        note=str(worksheet.cell_value(p,42))
                    #        label=str(worksheet.cell_value(p,41))
                    #        value=label.replace(" ","_")
                    #        value=value.lower()
                    #        cp["ArticleReach"]={ "note":note,"internal": True,"value": {"value": value,"label": label}} 

                    if worksheet.cell_value(p,43):
                            if fileN=="Macewan_licenses":
                                note=str(worksheet.cell_value(p,44))
                                value=int(worksheet.cell_value(p,43))
                            elif fileN=="Liverpool_licenses":
                                note=str(worksheet.cell_value(p,43))
                                value=0
                            cp["ConcurrentUsers"]={"note": note,"internal": True,"value": value,"type": {"id": "2c91808b74455016017566e8792f001f","name": "ConcurrentUsers","primary": True,"defaultInternal": True,"label": "Concurrent Users","description": "Specifies the number of allowed concurrent users","weight": 0,"type": "com.k_int.web.toolkit.custprops.types.CustomPropertyInteger"}}
                        
                    #if worksheet.cell_value(p,45):
                    #        label=""
                    #        value=""
                    #        note=""
                    #        note=str(worksheet.cell_value(p,46))
                    #        label=str(worksheet.cell_value(p,45))
                    #        value=label.replace(" ","_")
                    #        value=value.lower()
                    #        cp["CopyDigital"]={ "note":note,"internal": True,"value": {"value": value,"label": label}} 

                    #if worksheet.cell_value(p,47):
                    #        label=""
                    #        value=""
                    #        note=""
                    #        note=str(worksheet.cell_value(p,48))
                    #        label=str(worksheet.cell_value(p,47))
                    #        value=label.replace(" ","_")
                    #        value=value.lower()
                    #        cp["CopyPrint"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}    
 
                    if worksheet.cell_value(p,49):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,50))
                            label=str(worksheet.cell_value(p,49))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["CoursePackElectronic"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}    

                    if worksheet.cell_value(p,51):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,52))
                            label=str(worksheet.cell_value(p,51))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["CoursePackPrint"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}

                    #if worksheet.cell_value(p,53):
                    #        label=""
                    #        value=""
                    #        note=""
                    #        note=str(worksheet.cell_value(p,54))
                    #        label=str(worksheet.cell_value(p,53))
                    #        value=label.replace(" ","_")
                    #        value=value.lower()
                    #        cp["DistanceEducation"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}

                    #if worksheet.cell_value(p,55):
                    #        label=""
                    #        value=""
                    #        note=""
                    #        note=str(worksheet.cell_value(p,56))
                    #        label=str(worksheet.cell_value(p,55))
                    #        value=label.replace(" ","_")
                    #        value=value.lower()
                    #        cp["FairDealingClause"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}
 
                    #if worksheet.cell_value(p,57):
                    #        label=""
                    #        value=""
                    #        note=""
                    #        note=str(worksheet.cell_value(p,58))
                    #        label=str(worksheet.cell_value(p,57))
                    #        if label=="":
                    #            value="unspecified"
                    #            label="unspecified"
                    #        value=label.replace(" ","_")
                    #        value=value.lower()
                    #        cp["FairUseClause"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}
 
                    if worksheet.cell_value(p,59):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,60))
                            label=str(worksheet.cell_value(p,59))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["ILLElectronicSecure"]={ "note":note,"internal": True,"value": {"value": value,"label": label}} 

                        
                    if worksheet.cell_value(p,61):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,62))
                            label=str(worksheet.cell_value(p,61))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["ILLPrint"]={ "note":note,"internal": True,"value": {"value": value,"label": label}} 
  
                    if worksheet.cell_value(p,63):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,64))
                            label=str(worksheet.cell_value(p,63))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["ILLElectronicSecure"]={ "note":note,"internal": True,"value": {"value": value,"label": label}} 
  
                    #if worksheet.cell_value(p,65):
                    #        label=""
                    #        value=""
                    #        note=""
                    #        note=str(worksheet.cell_value(p,66))
                    #        label=str(worksheet.cell_value(p,65))
                    #        value=label.replace(" ","_")
                    #        value=value.lower()
                    #        cp["LinkElectronic"]={ "note":note,"internal": True,"value": {"value": value,"label": label}} 

                    #if worksheet.cell_value(p,67):
                    #        label=""
                    #        value=""
                    #        note=""
                    #        note=str(worksheet.cell_value(p,68))
                    #        label=str(worksheet.cell_value(p,67))
                    #        value=label.replace(" ","_")
                    #        value=value.lower()
                    #        cp["LMS"]={ "note":note,"internal": True,"value": {"value": value,"label": label}} 
 
                    if worksheet.cell_value(p,69):
                            pass
                            #OthersRestrictions
                    if worksheet.cell_value(p,70):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,71))
                            label=str(worksheet.cell_value(p,70))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["PerpetualAccess"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}

                       
                    #if worksheet.cell_value(p,72):
                    #        label=""
                    #        value=""
                    #        note=""
                    #        note=str(worksheet.cell_value(p,73))
                    #        label=str(worksheet.cell_value(p,72))
                    #        value=label.replace(" ","_")
                    #        value=value.lower()
                    #        cp["PersonsPerceptualDisabilities"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}
 
                       
                    if worksheet.cell_value(p,74):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,75))
                            label=str(worksheet.cell_value(p,74))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["ReservesElectronic"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}
 
                      
                    #if worksheet.cell_value(p,76):
                    #        label=""
                    #        value=""
                    #        note=""
                    #        note=str(worksheet.cell_value(p,77))
                    #        label=str(worksheet.cell_value(p,76))
                    #        value=label.replace(" ","_")
                    #        value=value.lower()
                    #        cp["ReservesPrint"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}

                       
                    #if worksheet.cell_value(p,78):
                    #        label=""
                    #        value=""
                    #        note=""
                    #        note=str(worksheet.cell_value(p,79))
                    #        label=str(worksheet.cell_value(p,78))
                    #        value=label.replace(" ","_")
                    #        value=value.lower()
                    #        cp["DepositRights"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}
 
                       
                    if worksheet.cell_value(p,80):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,81))
                            label=str(worksheet.cell_value(p,80))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["ScholarlySharing"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}
 
                    #if worksheet.cell_value(p,82):
                    #        label=""
                    #        value=""
                    #        note=""
                    #        note=str(worksheet.cell_value(p,83))
                    #        label=str(worksheet.cell_value(p,82))
                    #        value=label.replace(" ","_")
                    #        value=value.lower()
                    #        cp["TextDataMining"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}
  
                       
                    if worksheet.cell_value(p,84):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,85))
                            label=str(worksheet.cell_value(p,84))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["Walkins"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}
###################################################################
# WIDENER
#
#################################################

                elif orgL=="W":
                    cp={}
                    if worksheet.cell_value(p,5):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,6))
                            label=str(worksheet.cell_value(p,5))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["AllRightsReserved"]= {"note":note,"internal": True,"value": {"value": value,"label": label}} 
                        
                    if worksheet.cell_value(p,6):
                            note=""
                            value=str(worksheet.cell_value(p,6))
                            cp["AuthorizedUsers"]={"note": note, "internal": True,"value": value,"type": {"id": "2c91808574452945017565ac3f550004","name": "AuthorizedUsers","primary": True,"defaultInternal": True,"label": "Authorized User Definition","description": "Defines what constitutes an authorized user for the licensed resource.","weight": 0,"type": "com.k_int.web.toolkit.custprops.types.CustomPropertyText"}}
                        
                    if worksheet.cell_value(p,7):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,8))
                            label=str(worksheet.cell_value(p,7))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["AgreementConfidentiality"]= {"note":note,"internal": True,"value": {"value": value,"label": label}} 
                        

                    if worksheet.cell_value(p,9):
                            label=""
                            value=""
                            note=""
                            label=str(worksheet.cell_value(p,9))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["UserConfidentiality"]= {"internal": True,"value": {"value": value,"label": label}} 
                       
  
                    if worksheet.cell_value(p,11):
                            note=""
                            note=str(worksheet.cell_value(p,11))
                            cp["GovJurisdiction"]= {"internal": True,"value": note,"type": {"name": "GovJurisdiction","primary": True,"defaultInternal": True,"label": "Governing Jurisdiction","description": "Details the governing jurisdiction of the license agreement","weight": 0,"type": "com.k_int.web.toolkit.custprops.types.CustomPropertyText"}}
                        
                        
                    if worksheet.cell_value(p,12):
                            note=""
                            note=str(worksheet.cell_value(p,12))
                            cp["GovLaw"]= {"internal": True,"value": note,"type": {"name": "GovLaw","primary": True,"defaultInternal": True,"label": "Governing Law","description": "Details the governing law of the license agreement","weight": 0,"type": "com.k_int.web.toolkit.custprops.types.CustomPropertyText"}}
                       
                            #note=""
                    if worksheet.cell_value(p,13):
                            note=""
                            note=str(worksheet.cell_value(p,13))
                            cp["IndemLicensee"]={"note": note,"internal": True,"value": "Indemnification by Licensee","type": {"id": "2c91808b74455016017565beec470004","name": "IndemLicensee","primary": True,"defaultInternal": True,"label": "Indemnification by Licensee","description": "Indemnification by Licensee","weight": 0,"type": "com.k_int.web.toolkit.custprops.types.CustomPropertyText"}}
                        
                    if worksheet.cell_value(p,14):
                            note=str(worksheet.cell_value(p,14))
                            cp["IndemLicensor"] ={"note": note,"internal": True,"value": "Indemnification by Licensor","type": {"id": "2c91808b74455016017565bf87920005","name": "IndemLicensor","primary": True,"defaultInternal": True,"label": "Indemnification by Licensor","description": "Indemnification by Licensor","weight": 0,"type": "com.k_int.web.toolkit.custprops.types.CustomPropertyText"}}
                       
                            note=""
                    if worksheet.cell_value(p,15):
                            note=str(worksheet.cell_value(p,15))
                    if worksheet.cell_value(p,24):
                            label=""
                            value=""
                            note=""
                            label=str(worksheet.cell_value(p,24))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["SERU"]= {"internal": True,"value": {"value": value,"label": label}}
                       
                        
                    if worksheet.cell_value(p,35):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,36))
                            label=str(worksheet.cell_value(p,35))
                            value=str(worksheet.cell_value(p,36))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["AlumniAccess"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}
  
                    if worksheet.cell_value(p,37):
                            note=""
                            value=str(worksheet.cell_value(p,37))
                            cp["CopyrightLaw"]={"id": 529,"note": note,"internal": True,"value": value,"type": {"id": "2c91808574452945017565e389c2000a","name": "CopyrightLaw","primary": True,"defaultInternal": True,"label": "Applicable Copyright Law","description": "Specifies the copyright law applicable to the licensed resource","weight": 0,"type": "com.k_int.web.toolkit.custprops.types.CustomPropertyText"}}
 
                        
                    if worksheet.cell_value(p,38):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,39))
                            label=str(worksheet.cell_value(p,38))
                            value=str(worksheet.cell_value(p,38))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["ArchivingAllowed"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}
 
                    if worksheet.cell_value(p,41):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,42))
                            label=str(worksheet.cell_value(p,41))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["ArticleReach"]={ "note":note,"internal": True,"value": {"value": value,"label": label}} 

                    if worksheet.cell_value(p,43):
                            if fileN=="Macewan_licenses":
                                note=str(worksheet.cell_value(p,44))
                                value=int(worksheet.cell_value(p,43))
                            elif fileN=="Liverpool_licenses":
                                note=str(worksheet.cell_value(p,43))
                                value=0
                            cp["ConcurrentUsers"]={"note": note,"internal": True,"value": value,"type": {"id": "2c91808b74455016017566e8792f001f","name": "ConcurrentUsers","primary": True,"defaultInternal": True,"label": "Concurrent Users","description": "Specifies the number of allowed concurrent users","weight": 0,"type": "com.k_int.web.toolkit.custprops.types.CustomPropertyInteger"}}
                        
                    if worksheet.cell_value(p,45):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,46))
                            label=str(worksheet.cell_value(p,45))
                            value=str(worksheet.cell_value(p,45))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["CopyDigital"]={ "note":note,"internal": True,"value": {"value": value,"label": label}} 

                    if worksheet.cell_value(p,47):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,48))
                            label=str(worksheet.cell_value(p,47))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["CopyPrint"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}    
 
                    if worksheet.cell_value(p,49):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,50))
                            label=str(worksheet.cell_value(p,49))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["CoursePackElectronic"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}    

                    if worksheet.cell_value(p,51):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,52))
                            label=str(worksheet.cell_value(p,51))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["CoursePackPrint"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}

                    if worksheet.cell_value(p,53):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,54))
                            label=str(worksheet.cell_value(p,53))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["DistanceEducation"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}

                    if worksheet.cell_value(p,55):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,56))
                            label=str(worksheet.cell_value(p,55))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["FairDealingClause"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}
 
                    if worksheet.cell_value(p,57):
                            label=""
                            value=str(worksheet.cell_value(p,57))
                            note=""
                            note=str(worksheet.cell_value(p,58))
                            label=str(worksheet.cell_value(p,57))
                            #if label=="":
                            #    value="unspecified"
                            #    label="unspecified"
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["FairUseClause"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}
 
                    if worksheet.cell_value(p,59):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,60))
                            label=str(worksheet.cell_value(p,59))
                            value=str(worksheet.cell_value(p,59))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["ILLElectronicSecure"]={ "note":note,"internal": True,"value": {"value": value,"label": label}} 

                        
                    if worksheet.cell_value(p,61):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,62))
                            label=str(worksheet.cell_value(p,61))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["ILLPrint"]={ "note":note,"internal": True,"value": {"value": value,"label": label}} 
  
                    if worksheet.cell_value(p,63):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,64))
                            label=str(worksheet.cell_value(p,63))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["ILLElectronicSecure"]={ "note":note,"internal": True,"value": {"value": value,"label": label}} 
  
                    if worksheet.cell_value(p,65):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,66))
                            label=str(worksheet.cell_value(p,65))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["LinkElectronic"]={ "note":note,"internal": True,"value": {"value": value,"label": label}} 

                    if worksheet.cell_value(p,67):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,68))
                            label=str(worksheet.cell_value(p,67))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["LMS"]={ "note":note,"internal": True,"value": {"value": value,"label": label}} 
 
                    if worksheet.cell_value(p,69):
                            note=""
                            value=str(worksheet.cell_value(p,69))
                            cp["OtherRestrictions"]={"note": note,"internal": True,"value": value,"type": {"name": "OtherRestrictions","primary": True,"defaultInternal": True,"label": "Other Restrictions","description": "A blanket term to capture restrictions on a licensed resource not covered by established terms","weight": 0,"type": "com.k_int.web.toolkit.custprops.types.CustomPropertyText"}}
                            #OthersRestrictions
                    if worksheet.cell_value(p,70):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,71))
                            label=str(worksheet.cell_value(p,70))
                            value=str(worksheet.cell_value(p,70))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["PerpetualAccess"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}

                       
                    if worksheet.cell_value(p,72):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,73))
                            label=str(worksheet.cell_value(p,72))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["PersonsPerceptualDisabilities"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}
 
                       
                    if worksheet.cell_value(p,74):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,75))
                            label=str(worksheet.cell_value(p,74))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["ReservesElectronic"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}
 
                      
                    if worksheet.cell_value(p,76):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,77))
                            label=str(worksheet.cell_value(p,76))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["ReservesPrint"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}

                       
                    if worksheet.cell_value(p,78):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,79))
                            label=str(worksheet.cell_value(p,78))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["DepositRights"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}
 
                       
                    if worksheet.cell_value(p,80):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,81))
                            label=str(worksheet.cell_value(p,80))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["ScholarlySharing"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}
 
                    if worksheet.cell_value(p,82):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,83))
                            label=str(worksheet.cell_value(p,82))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["TextDataMining"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}
  
                       
                    if worksheet.cell_value(p,84):
                            label=""
                            value=""
                            note=""
                            note=str(worksheet.cell_value(p,85))
                            label=str(worksheet.cell_value(p,84))
                            value=label.replace(" ","_")
                            value=value.lower()
                            cp["Walkins"]={ "note":note,"internal": True,"value": {"value": value,"label": label}}
 

                org=licenses(uuidpol)
                #(self, ermName, ermOrgid, ermOrgName, ermDescription, aliases, fileName):
                org.licencePrint(ermName, vendor,ermDescriptionA, ERMlicType, ERMlicstatus, ermStartDate, ermendDate,docs, aliases, consortia,ermopenEnded,cp,path)


def notes(spreadsheet,org):
    wb = xlrd.open_workbook(spreadsheet)
    fileN=org
    worksheet = wb.sheet_by_name("all")
    print("no rows: ", worksheet.nrows)
    print("no columns: ", worksheet.ncols)
    count=0
    for p in range(worksheet.nrows):
        if (p!=0):
                #note 1
                print("\n######################## Record No."+str(p)+"######################\n")
                orgname=worksheet.cell_value(p,1)
                linkId=""
                if org=="Macewan_licenses":
                    linkId=get_licId_Macewan(orgname)
                elif org=="Liverpool_licenses":
                    linkId=get_licId_Liverpool(orgname)
                elif org=="Widener_licenses":
                    linkId=get_licId_Wineder(orgname)

                content=""
                title=""
                if worksheet.cell_value(p,19):
                    title= "License Note"
                    content= worksheet.cell_value(p,19)
                    #print_notes(title,linkId,cont,path,tn)
                    print_notes(title,linkId,content)
                #note 2
                content=""
                title=""
                if worksheet.cell_value(p,21):
                    title= "License Status Note: "+str(worksheet.cell_value(p,21))
                    content= worksheet.cell_value(p,21)
                    print_notes(title,linkId,content)
                #note 3
                content=""
                title=""
                if worksheet.cell_value(p,16):
                    days=str(worksheet.cell_value(p,16))
                    days=days.replace(".0","")
                    title= "License End Advance Notice Required ("+days+" Days)"
                    content= title
                    print_notes(title,linkId,content)
                #note 4
                content=""
                title=""
                if worksheet.cell_value(p,26):
                    title= "Reviewer Notes" 
                    content= str(worksheet.cell_value(p,26))
                    print_notes(title,linkId,content)
                #note 5
                content=""
                title=""
                if worksheet.cell_value(p,15):
                    title= "License Duration ("+str(worksheet.cell_value(p,15))+")" 
                    content= title
                    print_notes(title,linkId,content)

def notes_liverpool(spreadsheet,org):
    wb = xlrd.open_workbook(spreadsheet)
    fileN=org
    worksheet = wb.sheet_by_name("all")
    print("no rows: ", worksheet.nrows)
    print("no columns: ", worksheet.ncols)
    typenote="07948e43-e798-4156-aaf8-78b2c5b89a8f"
    count=0
    for p in range(worksheet.nrows):
        if (p!=0):
                #note 1
                print("\n######################## Record No."+str(p)+"######################\n")
                linkId=""
                content=""
                title=""
                orgname=worksheet.cell_value(p,1)
                if org=="Macewan_licenses":
                    linkId=get_licId_Macewan(orgname)
                elif org=="Liverpool_licenses":
                    linkId=get_licId_Liverpool(orgname)
                elif org=="Widener_licenses":
                    linkId=get_licId_Liverpool(orgname)
                
                if worksheet.cell_value(p,16):
                    days=str(worksheet.cell_value(p,16))
                    days=days.replace(".0","")
                    title= "License End Advance Notice Required "+days
                    content= "License End Advance Notice Required "+ days
                    print_notes(title,linkId,content)


if __name__ == "__main__":

    """This is the Starting point for the script"""    
    client="W"    
    if (client=="M"):
        file="Widener\licenses\Widener_License_SourceDataFinal.xlsx"
        ruta="Widener\licenses\widener_licenses.json"
        customer="Macewan_licenses"
        #readSpreadsheet(file,ruta,customer,"M")# "macewan\licenses\MacEwan_License_SourceDataFinal.xlsx", "Macewan_licenses")
        notes(file, customer)
    elif (client=="L"):
        file="liverpool\licenses\Liverpool_License_SourceDataFinal.xlsx"
        ruta="liverpool\licenses\liverpool_licenses.json"
        customer="Liverpool_licenses"
        readSpreadsheet(file,ruta,customer,"L")
        notes_liverpool(file,ruta,customer)
    elif (client=="W"):
        file="Widener\licenses\Widener_License_SourceDataFinal.xlsx"
        ruta="Widener\licenses\widener_licenses.json"
        customer="Widener_licenses"
        #readSpreadsheet(file,ruta,customer,"W")# "macewan\licenses\MacEwan_License_SourceDataFinal.xlsx", "Macewan_licenses")
        notes(file, customer)
    #readSpreadsheet("Liverpool\licenses\Liverpool_License_SourceDataFinal.xlsx", "Liverpool")
           