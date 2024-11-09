unit ulcsAccountViewUtility;

{This unit has methods and classes specific to AccountView accounting integration}

{-------------------------------------------------------------------------------
  Revision History:

  Date        Ref             Person      Comments
  2023/10/26  6.38.501/10058  PAB/DAK     Enh: Accounting integration with Visma AccountView
}

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, LResources, laz2_DOM, xmlutils, fpjson,
  IniFiles,
  ulcsAccUtility, uncsRESTClient, u_csLog,
  uncsAccountViewOAuthUtil, ulcsMessages, Math;

type

  { TAVParam }

  TAVParam = class(TAccParam)

  end;

  { TAVClient }

  TAVClient = class(TAccClient)
  private
    FAPIHeaders: TStringList;
    FAPIBaseURL: string;
    FClientId: string;
    FClientSecret: string;
    FRedirectURI: string;
    FCompanyCode: string;
    FPortNumber: integer;
    AccOA2Handler: TAccountViewOAuthHandler;
    FVATTranslations: TStringList;

    procedure DataReceived;
    procedure ReadConfigData;
    function GetAPIHeaders: TStringList;
    function SendToAccServer(AccResource: string; AccJsonObject: TJSONObject; BusinessObject: String): TJSONData;
    function PrepareAccContactTemplate(): TJSONObject;
    function GetFirst(URLEncodedODataFilter, BusinessObject: string): TJSONData;
    function GetFirstId(URLEncodedODataFilter, BusinessObject, Code: string): string;
    function PrepareAccInvoiceTemplate(): TJSONObject;
    function PrepareAccInvoiceData(InclVATX: string): TJSONObject;
    function PrepareAccInvoiceDetailData(LineTotalVATAmount, LineTotalSalePrice, LineVatCode: String): TJSONObject;
    function GetAccBTWTemplate(LineVatCode: String): String;

  public
    constructor Create;
    destructor Destroy; override;
    procedure Authenticate;
    {Creates contact in AccServer and returns the ContactID}
    function CreateContact(EulContactNode: TDOMNode; const LookupList: TStringList): string;
    {Creates invoice in AccServer and returns the InvoiceID}
    function CreateInvoice(EulInvoiceNode: TDOMNode; const LookupList: TStringList): string;
    function GetDebtorCodeId(BusinessObject: string): string;
    function GetDayBookCodeID(BusinessObject: string): string;
    function GetCompanyId : string;
  end;

const
  VATCodePrefix = 'V-';
  ContactPrefix = 'C-';
  DebtorCodePrefix = 'D-';
  DailyCodePrefix = 'DA-';
var
  CompanyId: string;

implementation

{ TAVClient }

constructor TAVClient.Create;
begin
  inherited;
  FAPIHeaders := nil;
  FVATTranslations := TStringList.Create;
  ReadConfigData;
end;

destructor TAVClient.Destroy;
begin
  inherited;
  if Assigned(FAPIHeaders) then
    FreeAndNil(FAPIHeaders);
  if Assigned(FVATTranslations) then
    FreeAndNil(FVATTranslations);
  if Assigned(AccOA2Handler) then
    FreeAndNil(AccOA2Handler);
end;

procedure TAVClient.Authenticate;
begin
  AccOA2Handler := TAccountViewOAuthHandler.Create(FClientId, FClientSecret, FRedirectURI, FPortNumber, @DataReceived);
  AccOA2Handler.InitiateOAuth;
  while not AccOA2Handler.CallBackReceived do
  ;
  if AccOA2Handler.AccessToken = string.Empty then
    raise Exception.Create(rsErrorAuthenticationFailed + ' [AccountView OAuth2] ' + AccOA2Handler.ErrorMessage);
end;

function TAVClient.CreateContact(EulContactNode: TDOMNode;const LookupList: TStringList): string;
var
  AccContact, JSONObject: TJSONObject;
  AccResponse, JSONData, JSONData1: TJSONData;
  JSONArray: TJSONArray;
begin
  try
    Result := string.Empty;
    AccContact := PrepareAccContactTemplate();
    FillTemplate(AccContact, EulContactNode, LookupList);
    AccResponse := SendToAccServer('/accountviewdata', AccContact,'AR1');

    if Assigned(AccResponse) then
      if TJSONString(AccResponse.FindPath('CONTACT')) <> nil then
      begin
        if AccResponse.JSONType = jtObject then
        begin
          JSONData := GetJSON(AccResponse.AsJSON);
          if JSONData is TJSONObject then
          begin
            JSONObject := TJSONObject(JSONData);
            JSONArray := JSONObject.Arrays['CONTACT'];
            JSONData1 := JSONArray.Items[0];
            if JSONData1 is TJSONObject then
              Result := TJSONObject(JSONData1).Get('SUB_NR');
          end;
        end;
      end
      else
      begin
        ELLog.WriteLog('JSONObject expected in the reponse, but found something else',
          ltError, 'TAVClient.CreateContact', 'After getting response from Server');
        ELLog.WriteLog(AccResponse.FormatJSON(), ltData, 'TAVClient.CreateContact',
          'After getting response from Server');
      end;
  finally
    if Assigned(AccContact) then
      FreeAndNil(AccContact);
    if Assigned(AccResponse) then
      FreeAndNil(AccResponse);
  end;
end;

function TAVClient.CreateInvoice(EulInvoiceNode: TDOMNode; const LookupList: TStringList): string;
var
  AccInvoice, jInvoiceTableDataObject, AccInvoiceData,
  AccInvoiceDetailDataValues, DetailDataObject: TJSONObject;
  JSONObject: TJSONObject;
  AccResponse, JSONData, JSONData1: TJSONData;
  LNode: TDomNode;
  RowsArray, JSONArray,
  DetailDataarray: TJSONArray;
  LTSPrice: double;
  LineTotalVATAmount, LineTotalSalePrice, LineVatCode: String;
begin
  try
    Result := string.Empty;
    AccInvoice := PrepareAccInvoiceTemplate();
    FillTemplate(AccInvoice, EulInvoiceNode, LookupList);
    jInvoiceTableDataObject := TJSONObject.Create;
    AccInvoiceData := PrepareAccInvoiceData(EulInvoiceNode.FindNode('InclVATX').TextContent);
    FillTemplate(AccInvoiceData, EulInvoiceNode, LookupList);
    jInvoiceTableDataObject.Add('Data', AccInvoiceData);
    RowsArray := TJSONArray.Create;
    LNode := EulInvoiceNode.FindNode('InvoiceLine');
    while (LNode <> nil) do
    begin
      {After InvoiceLine nodes, there may be other things like InvoiceRef, InvoiceVat}
      if LNode.NodeName = 'InvoiceLine' then
      begin
        Result := string.Empty;
        if LNode.FindNode('LineVATCode') <> nil then
          LineVatCode := LNode.FindNode('LineVATCode').TextContent;
        if LNode.FindNode('LineTotalVATAmount') <> nil then
          LineTotalVATAmount := LNode.FindNode('LineTotalVATAmount').TextContent
        else
          LineTotalVATAmount := '0';
         if LNode.FindNode('LineTotalSalePrice') <> nil then
          LineTotalSalePrice := LNode.FindNode('LineTotalSalePrice').TextContent
        else
          LineTotalSalePrice := '0';
        if (LNode.FindNode('LineSalePrice') <> nil) and (LNode.FindNode('LineQty') <> nil) and (LNode.FindNode('LineRef') <> nil) then
        begin
          AccInvoiceDetailDataValues := TJSONObject.Create;
          AccInvoiceDetailDataValues := PrepareAccInvoiceDetailData(LineTotalVATAmount, LineTotalSalePrice, LineVatCode);
          FillTemplate(AccInvoiceDetailDataValues, LNode, LookupList);
          RowsArray.Add(AccInvoiceDetailDataValues);
        end;
      end;
      LNode := LNode.NextSibling;
    end;
    DetailDataObject := TJSONObject.Create;
    DetailDataarray := TJSONArray.Create;
    DetailDataObject.Add('Rows', RowsArray);
    DetailDataarray.Add(DetailDataObject);
    jInvoiceTableDataObject.Add('DetailData', DetailDataarray);

    AccInvoice.Add('TableData', jInvoiceTableDataObject);
    AccResponse := SendToAccServer('/accountviewdata', AccInvoice, 'DJ2');

    if Assigned(AccResponse) then
      if TJSONString(AccResponse.FindPath('DJ_PAGE')) <> nil then
      begin
        if AccResponse.JSONType = jtObject then
        begin
          JSONData := GetJSON(AccResponse.AsJSON);
          if JSONData is TJSONObject then
          begin
            JSONObject := TJSONObject(JSONData);
            JSONArray := JSONObject.Arrays['DJ_PAGE'];
            JSONData1 := JSONArray.Items[0];
            if JSONData1 is TJSONObject then
              Result := TJSONObject(JSONData1).Get('INV_NR');
          end;
        end;
      end
      else
      begin
        ELLog.WriteLog('Create Invoice failed at AccountView', ltError, 'TAVClient.CreateInvoice',
           'After getting response from Server');
        ELLog.WriteLog(AccResponse.FormatJSON(), ltData, 'TAVClient.CreateInvoice', 'After getting response from Server');
      end
  finally
    if Assigned(AccInvoice) then
      FreeAndNil(AccInvoice);
    if Assigned(AccResponse) then
      FreeAndNil(AccResponse);
  end;
end;

function TAVClient.GetDebtorCodeId(BusinessObject: string): string;
begin
  Result := GetFirstId('/accountviewdata?BusinessObject=' + BusinessObject + '&PageSize=1&FilterControlSource1=ACCT_Name&FilterOperator1=Equal&FilterValue1=Debiteuren&FilterValueType1=C', BusinessObject, 'ACCT_NR');
end;

function TAVClient.GetDayBookCodeID(BusinessObject: string): string;
begin
  Result := GetFirstId('/accountviewdata?BusinessObject=' + BusinessObject + '&PageSize=1&FilterControlSource1=DJ_NAME&FilterOperator1=Equal&FilterValueType1=C&FilterValue1=Verkoopboek', BusinessObject, 'DJ_CODE');
end;

function TAVClient.GetCompanyId: string;
begin
  Result := GetFirstId('/companies','','Id');
end;

{------------------------------------------------------------------------------}
{---------------------------------Private methods------------------------------}
{------------------------------------------------------------------------------}

procedure TAVClient.DataReceived;
begin
  // Empty method is sufficient
end;

procedure TAVClient.ReadConfigData;
const
  constINIVATTranslationsSection = 'VATTranslations';
var
  IniFileName: string;
  ini: TIniFile;
  PortNumberString: string;
begin
  IniFileName := ChangeFileExt(ApplicationName, '.ini');
  if not FileExists(IniFileName) then
    raise Exception.Create(rsConfigurationFileMissing);
  ini := TIniFile.Create(IniFileName);
  try
    FAPIBaseURL := ini.ReadString('AccountViewAPI', 'APIBaseURL', '');
    FClientId := ini.ReadString('AccountViewAPI', 'client_id', '');
    FClientSecret := ini.ReadString('AccountViewAPI', 'client_secret', '');
    FCompanyCode := ini.ReadString('AccountViewAPI', 'x-companycode', '');
    FRedirectURI :=  ini.ReadString('AccountViewAPI', 'RedirectURI', '');
    PortNumberString := ini.Readstring('AccountViewAPI', 'PortNumber', '');
    FPortNumber := StrToIntDef(PortNumberString, FPortNumber);

    if (FClientId = string.Empty) or (FClientSecret = string.Empty) then
      raise Exception.CreateFmt(rsInvalidConfiguration, [rsInvalidToken]);
    if (FCompanyCode = string.Empty) then
      raise Exception.CreateFmt(rsInvalidConfiguration, ['x-companycode']);
    if (FRedirectURI = string.Empty) then
      raise Exception.CreateFmt(rsInvalidConfiguration, ['RedirectURI']);
    if (PortNumberString = string.Empty) then
      raise Exception.CreateFmt(rsInvalidConfiguration, ['PortNumber']);
    if (FAPIBaseURL = string.Empty) then
      raise Exception.CreateFmt(rsInvalidConfiguration, [rsInvalidAPIBaseURL]);
    if ini.SectionExists(constINIVATTranslationsSection) then
      ini.ReadSectionRaw(constINIVATTranslationsSection, FVATTranslations);


  finally
    if Assigned(ini) then
      FreeAndNil(ini);
  end;
end;

function TAVClient.GetAPIHeaders: TStringList;
begin
  Result := nil;

  {If first time, read auth info from the configuration}
  if not Assigned(FAPIHeaders) then
    FAPIHeaders := TStringList.Create
  else
    FAPIHeaders.Clear;
  TEuRESTClient.AddHeader(FAPIHeaders, 'Client-Id', FClientId);
  TEuRESTClient.AddHeader(FAPIHeaders, 'x-company', CompanyId);
  TEuRESTClient.AddHeader(FAPIHeaders, 'Authorization','Bearer ' + AccOA2Handler.AccessToken);
  Result := FAPIHeaders;
end;

function TAVClient.PrepareAccContactTemplate(): TJSONObject;
var
  AccContact: TJSONObject;
  Jobject, JObject1, JObject2, JObject3, JObject4, JObject5: TJSONObject;
  JArray, JArray1, jArray2: TJSONArray;
  Name: String;
  i: Integer;
begin
  Result := nil;
  AccContact := TJSONObject.Create;
  with AccContact do
  begin
    Add('BookDate', '"2023-9-7T10:39:05.276Z"');
    Add('BusinessObject', '"AR1"');

    JObject := TJSONObject.Create;
    JObject1 := TJSONObject.Create;
    JObject1.Add('Name', '"CONTACT"');

    JArray := TJSONArray.Create;

    for i := 0 to 12 do
    begin
      case i of
        0: Name := 'RowId';
        1: Name := 'SUB_NR';
        2: Name := 'ACCT_NR';
        3: Name := 'SRC_CODE';
        4: Name := 'ACCT_NAME';
        5: Name := 'ADDRESS1';
        6: Name := 'POST_CODE';
        7: Name := 'CITY';
        8: Name := 'CNT_CODE';
        9: Name := 'TEL_BUS';
        10: Name := 'TEL_MOB';
        11: Name := 'MAIL_BUS';
        12: Name := 'WWW_URL';
      end;

      JObject2 := TJSONObject.Create;
      JObject2.Add('Name', '"' + Name + '"');
      JObject2.Add('FieldType', '"C"');
      JArray.Add(JObject2);
    end;

    JObject1.Add('Fields', JArray);
    JObject.Add('Definition', JObject1);
    Add('Table', JObject);

    JObject3 := TJSONObject.Create;
    JObject4 := TJSONObject.Create;
    JObject5 := TJSONObject.Create;
    JObject4.Add('Data', JObject3);
    JArray1 := TJSONArray.Create;
    JArray2 := TJSONArray.Create;
    JArray1.Add(JObject5);

    JArray2.Add('"1"');
    JArray2.Add('[Ref]');
    JArray2.Add('{' + DebtorCodePrefix + '"GL1"}');
    JArray2.Add('[Ref]');
    JArray2.Add('<CombinedName>');
    JArray2.Add('<CombinedAddress>');
    JArray2.Add('Zip');
    JArray2.Add('Town');
    JArray2.Add('CountryCode');
    JArray2.Add('WorkTel');
    JArray2.Add('MobTel');
    JArray2.Add('Email');
    JArray2.Add('Web');
    JObject5.Add('Values', JArray2);
    JObject3.Add('Rows', JArray1);

    Add('TableData', JObject4);
  end;
  Result := AccContact;
end;

function TAVClient.GetFirstId(URLEncodedODataFilter, BusinessObject,
  Code: string): string;
var
  Response: TJSONData;
  JObject: TJSONObject;
begin
  Result := string.Empty;
  Response := nil;
  JObject := nil;
  Response := GetFirst(URLEncodedODataFilter, BusinessObject);
  if Assigned(Response) then
  begin
    JObject := TJSONObject(Response);
    Result := TJSONString(JObject.FindPath(Code)).Value;
  end;
end;

function TAVClient.PrepareAccInvoiceTemplate(): TJSONObject;
var
  AccInvoice: TJSONObject;
  JArray, JArray1, JArray2: TJSONArray;
  JObject, JObject1, JObject2, JObject3, JObject4: TJSONObject;
  i: integer;
  Name,FieldType : String;
begin
  Result := nil;
  AccInvoice := TJSONObject.Create;
  JArray := TJSONArray.Create;

  with AccInvoice do
  begin
   Add('BookDate', '"2023-9-7T10:39:05.276Z"');
    Add('BusinessObject', '"DJ2"');

    JObject := TJSONObject.Create;
    JObject1 := TJSONObject.Create;
    JObject1.Add('Name', '"DJ_PAGE"');

    JArray := TJSONArray.Create;

    for i := 0 to 11 do
    begin
      case i of
        0: begin
             Name := 'RowId';
             FieldType := 'C';
           end;
        1: begin
             Name := 'DJ_CODE';
             FieldType := 'C';
           end;
        2: begin
             Name := 'SUB_NR';
             FieldType := 'C';
           end;
        3: begin
             Name := 'INV_NR';
             FieldType := 'C';
           end;
        4: begin
             Name := 'TRN_DATE';
             FieldType := 'T';
           end;
        5: begin
             Name := 'PERIOD';
             FieldType := 'N';
           end;
        6: begin
             Name := 'HDR_DESC';
             FieldType := 'C';
           end;
        7: begin
             Name := 'GROSS_FLG';
             FieldType := 'L';
           end;
        8: begin
             Name := 'CUR_CODE';
             FieldType := 'C';
           end;
        9: begin
             Name := 'CRED_DAYS';
             FieldType := 'N';
           end;
        10: begin
             Name := 'PD_PCT';
             FieldType := 'N';
           end;
        11: begin
              Name := 'COMMENT1';
              FieldType := 'C';
            end;
      end;
      JObject2 := TJSONObject.Create;
      JObject2.Add('Name', '"' + Name + '"');
      JObject2.Add('FieldType', '"' + FieldType + '"');
      JArray.Add(JObject2);
    end;
    JObject1.Add('Fields', JArray);
    JObject.Add('Definition', JObject1);

    jArray2 := TJSONArray.Create;
    JArray1 := TJSONArray.Create;
    JObject3 := TJSONObject.Create;
    JObject3.Add('Name', '"DJ_LINE"');

    for i := 0 to 7 do
    begin
      case i of
        0: begin
          Name := 'RowId';
          FieldType := 'C';
        end;
        1: begin
          Name := 'HeaderId';
          FieldType := 'C';
        end;
        2: begin
          Name := 'ACCT_NR';
          FieldType := 'C';
        end;
        3: begin
          Name := 'TRN_DESC';
          FieldType := 'C';
        end;
        4: begin
          Name := 'AMOUNT';
          FieldType := 'N';
        end;

        5: begin
          Name := 'VAT_Code';
          FieldType := 'C';
        end;
        6: begin
          Name := 'VAT_AMT';
          FieldType := 'N';
        end;
        7: begin
          Name := 'TRN_QTY';
          FieldType := 'N';
        end;
      end;
      JObject4 := TJSONObject.Create;
      JObject4.Add('Name', '"' + Name + '"');
      JObject4.Add('FieldType', '"' + FieldType + '"');
      JArray1.Add(JObject4);
    end;
    JObject3.Add('Fields', JArray1);
    jArray2.Add(JObject3);

    JObject.Add('DetailDefinitions', jArray2);
    Add('Table', JObject);
  end;
  Result := AccInvoice;
end;

function TAVClient.PrepareAccInvoiceData(InclVATX: string): TJSONObject;
var
  DataObject, ValuesObject : TJSONObject;
  RowsArray, ValuesArray : TJSONArray;
begin
  Result := nil;
  DataObject := TJSONObject.Create;
  ValuesObject := TJSONObject.Create;
  ValuesArray := TJSONArray.Create;
  RowsArray := TJSONArray.Create;
  ValuesArray.Add('"1"');
  ValuesArray.Add('{' + DailyCodePrefix + '"DJ1"}');
  ValuesArray.Add('ContRef');
  ValuesArray.Add('[MainNumber]');
  ValuesArray.Add('RegDate');
  ValuesArray.Add('<Month>');
  ValuesArray.Add('Description');
  if (InclVATX = '1') then
   ValuesArray.Add('"TRUE"')
  else
    ValuesArray.Add('"False"');
  ValuesArray.Add('"EUR"');
  ValuesArray.Add('CreditCode');
  ValuesArray.Add('<AVDiscount>');
  ValuesArray.Add('InternalMemo');
  ValuesObject.Add('Values', ValuesArray);
  RowsArray.Add(ValuesObject);
  DataObject.Add('Rows', RowsArray);
  Result:=  DataObject;
end;

function TAVClient.PrepareAccInvoiceDetailData(LineTotalVATAmount, LineTotalSalePrice, LineVatCode: String): TJSONObject;
var
  ValuesObject :TJSONObject;
  ValuesArray: TJSONArray;
begin
  Result := nil;
  ValuesObject := TJSONObject.Create;
  ValuesArray := TJSONArray.Create;

  ValuesArray.Add('LineKeyNumber');
  ValuesArray.Add('"1"');
  ValuesArray.Add('LineNominalCode');
  ValuesArray.Add('LineDescription');
  ValuesArray.Add('"' + LineTotalSalePrice + '"');
 { ValuesArray.Add('LineVATCode')}
  ValuesArray.Add(GetAccBTWTemplate(LineVatCode));
  ValuesArray.Add('"' + LineTotalVATAmount + '"');
  ValuesArray.Add('LineQty');
  ValuesObject.Add('Values', ValuesArray);
  Result:= ValuesObject;
end;

function TAVClient.GetAccBTWTemplate(LineVatCode: String): String;
var
  VatCode, ValueString: string;
  i, PosEquals: integer;
begin
  VatCode := LineVatCode;
  for i := 0 to FVATTranslations.Count - 1 do
  begin
    PosEquals := Pos('=', FVATTranslations[i]); // Find the position of '='
    if (PosEquals > 0) and (Copy(FVATTranslations[i], 1, PosEquals - 1) = VatCode) then
    begin
      ValueString := Copy(FVATTranslations[i], PosEquals + 1, Length(FVATTranslations[i]) - PosEquals);
      Break;
    end;
  end;
end;

function TAVClient.GetFirst(URLEncodedODataFilter, BusinessObject: string): TJSONData;
var
  rc: TEuRESTClient;
  APIHeaders: TStringList;
  ResponseCode: integer;
  Response: TStringStream;
  ResponseJson: TJSONData;
  AccGetArray: TJSONArray;
  AccGetObject, AccGetObject1: TJSONObject;
  ReturnArray,AccGetArray1: TJSONArray;
  i: integer;
  ACCObjectName: string;
begin
  Result := nil;
  Response := nil;
  rc := TEuRESTClient.Create(FAPIBaseURL);
  try
    if (BusinessObject = 'GL1') then
    begin
      ACCObjectName := 'LEDGER';
    end
    else if (BusinessObject = 'AR1') then
    begin
      ACCObjectName := 'CONTACT';
    end
    else if (BusinessObject = 'DJ1') then
    begin
      ACCObjectName := 'DAILY';
    end
    else if (BusinessObject = 'DJ2') then
    begin
      ACCObjectName := 'DJ_PAGE';
    end;
    APIHeaders := GetAPIHeaders();
    Response := rc.Get(URLEncodedODataFilter, string.Empty, APIHeaders);
    ResponseCode := rc.ResponseStatusCode;
    if (ResponseCode <> 200) and (ResponseCode <> 404) then
      raise Exception.CreateFmt(rsInternalError, ['REST error in the get-method to check whether the resource exists (ResponseCode: ' + IntToStr(ResponseCode) + ')']);
    AccGetArray := nil;
    AccGetObject := nil;
    if ResponseCode = 200 then
    begin
      ResponseJson := GetJSON(Response.DataString);
      if ResponseJson is TJSONObject then
      begin
        AccGetObject := TJSONObject(ResponseJson);
      end
      else if ResponseJson is TJSONArray then
      begin
        AccGetArray := TJSONArray(ResponseJson);
      end;
    end;
    if Assigned(AccGetObject) then
    begin
      ReturnArray := AccGetObject.Get(ACCObjectName, AccGetArray);
      if ReturnArray <> nil then
      begin
        AccGetObject1 := TJSONObject(GetJSON(ReturnArray.AsJSON));
        for i := 0 to AccGetObject1.Count - 1 do
        begin
          Result := AccGetObject1.Items[i];
        end;
      end;
    end
    else if Assigned(AccGetArray) then
    begin
      AccGetArray1 := TJSONArray(GetJSON(AccGetArray.AsJSON));
      for i := 0 to AccGetArray1.Count - 1 do
      begin
        if AccGetArray1.Objects[i].Find('Code').AsString = FCompanyCode then
          Result := AccGetArray1.Objects[i];
      end;
    end;
  finally
    if Assigned(Response) then
      FreeAndNil(Response);
    if Assigned(rc) then
      FreeAndNil(rc);
  end;
end;

function TAVClient.SendToAccServer(AccResource: string; AccJsonObject: TJSONObject; BusinessObject: String): TJSONData;
var
  rc: TEuRESTClient;
  APIHeaders: TStringList;
  ResponseCode: integer;
  Response: TStringStream;
  Filter: string;
  AccCodeColumnName: string;
  AccCodeColumnValue: TJSONVariant;
  AccId: string;
  Index, i: integer;
  TableData, Data, Rows, Definition, Field: TJsonObject;
  RowsArray, ValuesArray, FieldsArray: TJsonArray;
begin
  Result := nil;
  Response := nil;
  rc := TEuRESTClient.Create(FAPIBaseURL);
  try
    ELLog.WriteLog('Request object:' + AccJsonObject.FormatJSON(), ltData, 'TAVClient.SendToAccServer', 'Beginning of try block');
    if (BusinessObject = 'AR1') then
    begin
      AccCodeColumnName := 'SUB_NR';
    end;
    if (BusinessObject = 'DJ2') then
    begin
      AccCodeColumnName := 'INV_NR';
    end;
    if AccJsonObject <> nil then
    begin
      TableData := AccJsonObject.Objects['TableData'];
      Data := TableData.Objects['Data'];
      RowsArray := Data.Arrays['Rows'];
      Rows := RowsArray.Objects[0];
      ValuesArray := Rows.Arrays['Values'];
      Definition := AccJsonObject.Objects['Table'].Objects['Definition'];
      FieldsArray := Definition.Arrays['Fields'];
      for i := 0 to FieldsArray.Count - 1 do
      begin
        Field := FieldsArray.Objects[i];
        if Field.Find('Name').Value = AccCodeColumnName then
        begin
          Index := i;
          Break;
        end;
      end;
      AccCodeColumnValue := ValuesArray.Items[Index].Value;
    end;
    if ((AccCodeColumnValue = Null) or (AccCodeColumnValue = nil)) then
    begin
      ELLog.WriteLog('Unable to export ' + AccResource + ' to AccountsView as' + ' ' + AccCodeColumnName + ' ' + 'is not available in the Eux',
           ltInfo, 'TAVClient.SendToAccServer', 'without AccCodeColumnValue - before posting');
    end
    else
    begin
      APIHeaders := GetAPIHeaders();
      Filter := Format('?BusinessObject=' + BusinessObject + '&PageSize=1' + '&FilterControlSource1=' + AccCodeColumnName + '&FilterValue1=' + AccCodeColumnValue + '&FilterValueType1=C&FilterOperator1=Equal', []);
      AccId := string.Empty;
      AccId := GetFirstId(AccResource + Filter, BusinessObject, AccCodeColumnName);
      if (AccId = AccCodeColumnValue) then
      begin
        if (AccId <> string.Empty) and (BusinessObject <> 'DJ2') then
        begin
          {Invoice already exists on the AccountView}
          ELLog.WriteLog(AccId + ' found, so using PUT method to update', ltInfo, 'TAVClient.SendToAccServer', 'with AccId - before posting');
          AccJsonObject.Add('id', AccId);
          ELLog.WriteLog(AccJsonObject.FormatJSON(), ltData, 'TAVClient.SendToAccServer', 'with AccId - before posting');
          Response := nil;
          Response := rc.Put(AccResource, AccJsonObject.AsJSON, APIHeaders);
        end
        else if (AccId <> string.Empty) and (BusinessObject = 'DJ2') then
        begin
        {There is no put functionality available in Accountview so that need to stop create same invoice multiple times
          and it sends 'id' of invoice which is already in AccountView}
          ELLog.WriteLog(AccId + ' found, but unable to update as PUT method is not available for invoice', ltInfo, 'TAVClient.SendToAccServer', 'with AccId - before posting');
          AccJsonObject.Add('id', AccId);
          ELLog.WriteLog('Invoice already exists at AccountView' + ' (' + AccId + ')', ltInfo, 'TAVClient.SendToAccServer', 'with AccId - before posting');
          Exit;
        end;
      end
      else
      begin
        {Invoice does not exist on the AccountView}
        ELLog.WriteLog(AccId + ' not found, so using POST method to create', ltInfo, 'TAVClient.SendToAccServer', 'without AccId - before posting');
        Response := nil;
        Response := rc.Post(AccResource, AccJsonObject.AsJSON, APIHeaders);
      end;
      ResponseCode := rc.ResponseStatusCode;
      Result := GetJSON(Response.DataString);
      ELLog.WriteLog('Response code:' + IntToStr(ResponseCode), ltInfo, 'TAVClient.SendToAccServer', 'With Response code - After Posting');
      ELLog.WriteLog('Response:' + Response.DataString, ltData, 'TAVClient.SendToAccServer', 'With Response - After Posting');
    end;
  finally
    if Assigned(Response) then
      FreeAndNil(Response);
    if Assigned(rc) then
      FreeAndNil(rc);
  end;
end;

end.
