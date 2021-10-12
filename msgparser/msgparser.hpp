#include "common.h"

std::string
wstr2str(wstring wstr, UINT CodePage)
{
    string result;
    PCHAR temp = NULL;
    INT Len = 0;

    Len = WideCharToMultiByte(CodePage, 0, wstr.c_str(), -1, NULL, NULL, NULL, NULL);
    if (Len == 0)
    {
        return result;
    }
    temp = new char[Len + 1];
    if (temp == NULL)
    {
        return result;
    }
    WideCharToMultiByte(CodePage, 0, wstr.c_str(), -1, temp, Len, NULL, NULL);
    temp[Len] = '\0';
    result.append(temp);
    delete[] temp;
    return result;
}

std::wstring
str2wstr(std::string str, UINT CodePage)
{
    INT          nLen = 0;
    PWCHAR       temp = NULL;
    std::wstring wstr;

    nLen = MultiByteToWideChar(CodePage, 0, str.c_str(), -1, NULL, NULL);
    if (nLen == 0)
    {
        return wstr;
    }
    temp = new WCHAR[nLen + 1];
    if (temp == NULL)
    {
        return wstr;
    }
    temp[nLen] = L'\0';
    nLen = MultiByteToWideChar(CodePage, 0, str.c_str(), -1, temp, nLen);
    wstr = temp;
    delete[]temp;
    return wstr;
}

class OutlookStorage
{
public:
    // attachment constants
    const string ATTACH_STORAGE_PREFIX = "__attach_version1.0_#";
    const string PR_ATTACH_FILENAME = "3704";
    const string PR_ATTACH_LONG_FILENAME = "3707";
    const string PR_ATTACH_DATA = "3701";
    const string PR_ATTACH_METHOD = "3705";
    const string PR_RENDERING_POSITION = "370B";
    const string PR_ATTACH_CONTENT_ID = "3712";
    const INT ATTACH_BY_VALUE = 1;
    const INT ATTACH_EMBEDDED_MSG = 5;
    // recipient constants
    const string RECIP_STORAGE_PREFIX = "__recip_version1.0_#";
    const string PR_DISPLAY_NAME = "3001";
    const string PR_EMAIL = "39FE";
    const string PR_EMAIL_2 = "403E";
    const string PR_RECIPIENT_TYPE = "0C15";
    const INT MAPI_TO = 1;
    const INT MAPI_CC = 2;
    // msg constants
    const string PR_SUBJECT = "0037";
    const string PR_BODY = "1000";
    const string PR_RTF_COMPRESSED = "1009";
    const string PR_SENDER_NAME = "0C1A";
    // property stream constants
    const string PROPERTIES_STREAM = "__properties_version1.0";
    const INT PROPERTIES_STREAM_HEADER_TOP = 32;
    const INT PROPERTIES_STREAM_HEADER_EMBEDED = 24;
    const INT PROPERTIES_STREAM_HEADER_ATTACH_OR_RECIP = 8;
    // name id storage name in root storage
    const string NAMEID_STORAGE = "__nameid_version1.0";
    const INT PT_UNSPECIFIED = 0;
    const INT PT_NULL = 1;
    const INT PT_I2 = 2;
    const INT PT_LONG = 3;
    const INT PT_R4 = 4;
    const INT PT_DOUBLE = 5;
    const INT PT_CURRENCY = 6;
    const INT PT_APPTIME = 7;
    const INT PT_ERROR = 10;
    const INT PT_BOOLEAN = 11;
    const INT PT_OBJECT = 13;
    const INT PT_I8 = 20;
    const INT PT_STRING8 = 30;
    const INT PT_UNICODE = 31;
    const INT PT_SYSTIME = 64;
    const INT PT_CLSID = 72;
    const INT PT_BINARY = 258;

    wstring Path;
    IStorage* Storage;
    map<LPOLESTR, STATSTG> subStorageStatistics;
    map<LPOLESTR, STATSTG> streamStatistics;
    INT propHeaderSize = PROPERTIES_STREAM_HEADER_TOP;
    OutlookStorage* parentMessage = NULL;
    OutlookStorage* TopParent = this->getTopParent();
    OutlookStorage* getTopParent()
    {
        if (this->parentMessage != NULL)
        {
            return this->parentMessage->TopParent;
        }
        return this;
    }

    BOOL IsTopParent()
    {
        if (this->parentMessage == NULL)
        {
            return TRUE;
        }
        return FALSE;
    }

    OutlookStorage(wstring storageFilePath)
    {
        IStorage* fileStorage = NULL;
        if (StgIsStorageFile(storageFilePath.c_str()) != 0)
        {
            return;
        }
        StgOpenStorage(storageFilePath.c_str(), NULL, STGM_READ | STGM_SHARE_DENY_WRITE, NULL, 0, &fileStorage);
        this->LoadStorage(fileStorage);
    }

    OutlookStorage(IStream* storageStream)
    {
        IStorage* memoryStorage = NULL;
        ILockBytes* memoryStorageBytes = NULL;
        PCHAR Buffer = NULL;
        ULONG Length = 0;
        _ULARGE_INTEGER v1 = { 0 };

        Buffer = (PCHAR)malloc(0x1000);
        RtlSecureZeroMemory(Buffer, 0x1000);
        storageStream->Read(Buffer, 0, &Length);
        CreateILockBytesOnHGlobal(NULL, TRUE, &memoryStorageBytes);
        memoryStorageBytes->WriteAt(v1, Buffer, Length, NULL);
        if (StgIsStorageILockBytes(memoryStorageBytes) != 0)
        {
            return;
        }
        StgOpenStorageOnILockBytes(memoryStorageBytes, NULL, STGM_READ | STGM_SHARE_DENY_WRITE, NULL, 0, &memoryStorage);
        this->LoadStorage(memoryStorage);

        if (Buffer != NULL)
        {
            free(Buffer);
            Buffer = NULL;
        }
        if (memoryStorage != NULL)
        {
            memoryStorage->Release();
        }
        if (memoryStorageBytes != NULL)
        {
            memoryStorageBytes->Release();
        }
    }

    OutlookStorage(IStorage* storage)
    {
        this->LoadStorage(storage);
    }

    ~OutlookStorage()
    {
        map<LPOLESTR, STATSTG>().swap(subStorageStatistics);
        map<LPOLESTR, STATSTG>().swap(streamStatistics);
        this->Storage->Release();
    }

    virtual void LoadStorage(IStorage* storage)
    {
        this->Storage = storage;
        IEnumSTATSTG* storageElementEnum = NULL;
        storage->EnumElements(0, NULL, 0, &storageElementEnum);

        while (TRUE)
        {
            ULONG elementStatCount;
            STATSTG elementStats[1] = { 0 };
            storageElementEnum->Next(1, elementStats, &elementStatCount);
            if (elementStatCount != 1)
            {
                break;
            }
            STATSTG elementStat = elementStats[0];
            switch (elementStat.type)
            {
            case 1:
                subStorageStatistics[elementStat.pwcsName] = elementStat;
                break;

            case 2:
                streamStatistics[elementStat.pwcsName] = elementStat;
                break;
            }
        }
        if (storageElementEnum != NULL)
        {
            storageElementEnum->Release();
        }
    }

    IStorage* CloneStorage(IStorage* source, BOOL closeSource)
    {
        IStorage* memoryStorage = NULL;
        ILockBytes* memoryStorageBytes = NULL;

        CreateILockBytesOnHGlobal(NULL, TRUE, &memoryStorageBytes);
        StgCreateDocfileOnILockBytes(memoryStorageBytes, STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE, 0, &memoryStorage);
        source->CopyTo(0, NULL, NULL, memoryStorage);
        memoryStorageBytes->Flush();
        memoryStorage->Commit(0);

        if (memoryStorage != NULL)
        {
            memoryStorage->Release();
        }
        if (memoryStorageBytes != NULL)
        {
            memoryStorageBytes->Release();
        }
        if (closeSource)
        {
            source->Release();
        }
        return memoryStorage;
    }

    BYTE* GetStreamBytes(string streamName, ULONG* length = NULL)
    {
        STATSTG streamStatStg;
        BYTE* iStreamContent = NULL;
        IStream* stream = NULL;

        for (map<LPOLESTR, STATSTG>::iterator i = this->streamStatistics.begin(); i != streamStatistics.end(); i++)
        {
            if (i->first == str2wstr(streamName, CP_UTF8))
            {
                streamStatStg = i->second;
                break;
            }
        }
        this->Storage->OpenStream(streamStatStg.pwcsName, NULL, STGM_READ | STGM_SHARE_EXCLUSIVE, 0, &stream);
        iStreamContent = (BYTE*)malloc(streamStatStg.cbSize.LowPart + 2);
        RtlSecureZeroMemory(iStreamContent, streamStatStg.cbSize.LowPart);
        stream->Read(iStreamContent, streamStatStg.cbSize.LowPart, length);
        if (stream != NULL)
        {
            stream->Release();
        }
        return iStreamContent;
    }

    wstring GetStreamAsString(string streamName)
    {
        ULONG Length = 0;
        wstring str;
        BYTE* StreamBytes = this->GetStreamBytes(streamName, &Length);
        PWCHAR v1 = (PWCHAR)StreamBytes;
        v1[Length / sizeof(WCHAR)] = '\0';
        str.append(v1, Length);
        free(StreamBytes);
        return str;
    }

    tuple<PVOID, wstring, ULONG> GetMapiPropertyFromStreamOrStorage(string propIdentifier)
    {
        tuple<PVOID, wstring, ULONG> Result;
        vector<LPOLESTR> propKeys;
        map<LPOLESTR, STATSTG>::iterator i;
        string propTag;
        int propType = PT_UNSPECIFIED;

        for (i = this->streamStatistics.begin(); i != streamStatistics.end(); i++)
        {
            propKeys.push_back(i->first);
        }
        for (i = this->subStorageStatistics.begin(); i != subStorageStatistics.end(); i++)
        {
            propKeys.push_back(i->first);
        }
        for (vector<LPOLESTR>::iterator i = propKeys.begin(); i != propKeys.end(); i++)
        {
            wstring str = *i;
            if (str.find(str2wstr("__substg1.0_" + propIdentifier, CP_UTF8)) == 0)
            {
                string temp = wstr2str(str, CP_UTF8);
                propTag = temp.substr(12, 8);
                propType = stoi(temp.substr(16, 4), 0, 0x10);
                break;
            }
        }
        string containerName = "__substg1.0_" + propTag;
        switch (propType)
        {
        case 0:    // PT_UNSPECIFIED
        {
            Result = make_tuple((PVOID)NULL, L"", 0);
            return Result;
        }
        case 30:    // PT_STRING8
        {
            Result = make_tuple((PVOID)NULL, this->GetStreamAsString(containerName), 0);
            return Result;
        }
        case 31:    // PT_UNICODE
        {
            Result = make_tuple((PVOID)NULL, this->GetStreamAsString(containerName), 0);
            return Result;
        }
        case 258:   // PT_BINARY
        {
            ULONG Length = 0;
            Result = make_tuple(this->GetStreamBytes(containerName, &Length), L"", Length);
            return Result;
        }
        case 13:    // PT_OBJECT
        {
            IStorage* v1;
            this->Storage->OpenStorage(str2wstr(containerName, CP_UTF8).c_str(), NULL, STGM_READ | STGM_SHARE_EXCLUSIVE, NULL, 0, &v1);
            Result = make_tuple(CloneStorage(v1, TRUE), L"", 0);
            return Result;
        }
        default:
        {
            Result = make_tuple((PVOID)NULL, L"", 0);
            return Result;
        }
        }
    }

    wstring GetMapiPropertyString(string propIdentifier)
    {
        PVOID v1;
        wstring v2;
        ULONG v3;
        tie(v1, v2, v3) = this->GetMapiPropertyFromStreamOrStorage(propIdentifier);
        return v2;
    }

    string HexArrayToString(PCHAR data, INT len)
    {
        const std::string hexme = "0123456789ABCDEF";
        std::string ret = "";
        for (int i = 0; i < len; i++)
        {
            ret.push_back(hexme[(data[i] & 0xF0) >> 4]);
            ret.push_back(hexme[data[i] & 0x0F]);
        }
        return ret;
    }

    INT GetMapiPropertyFromPropertyStream(string propIdentifier)
    {
        BYTE Flag = 0;
        map<LPOLESTR, STATSTG>::iterator i;
        for (i = this->streamStatistics.begin(); i != streamStatistics.end(); i++)
        {
            if (i->first == str2wstr(PROPERTIES_STREAM, CP_UTF8))
            {
                Flag = 1;
                break;
            }
        }
        if (Flag == 0)
        {
            return NULL;
        }
        ULONG propLength = 0;
        BYTE* propBytes = this->GetStreamBytes(PROPERTIES_STREAM, &propLength);
        for (ULONG i = this->propHeaderSize; i < propLength; i = i + 16)
        {
            USHORT propType = *((USHORT*)(propBytes + i));
            CHAR propIdent[] = { (CHAR)propBytes[i + 3] , (CHAR)propBytes[i + 2] };
            string propIdentString = HexArrayToString(propIdent, 2);
            if (propIdentString != propIdentifier)
            {
                continue;
            }
            switch (propType)
            {
            case 2:    // PT_I2
                return *((INT16*)(propBytes + i + 8));

            case 3:    // PT_LONG
                return *((INT32*)(propBytes + i + 8));

            default:
            {
                return NULL;
            }
            }
        }
        return NULL;
    }

    INT GetMapiPropertyInt16(string propIdentifier)
    {
        return this->GetMapiPropertyFromPropertyStream(propIdentifier);
    }

    INT GetMapiPropertyInt32(string propIdentifier)
    {
        return this->GetMapiPropertyFromPropertyStream(propIdentifier);
    }

    BYTE* GetMapiPropertyBytes(string propIdentifier, ULONG* Length)
    {
        PVOID v1;
        wstring v2;
        tie(v1, v2, *Length) = this->GetMapiPropertyFromStreamOrStorage(propIdentifier);
        return (BYTE*)v1;
    }

    IStorage* GetMapiPropertyIStorage(string propIdentifier)
    {
        PVOID v1;
        wstring v2;
        ULONG v3;
        tie(v1, v2, v3) = this->GetMapiPropertyFromStreamOrStorage(propIdentifier);
        return (IStorage*)v1;
    }
};

class Attachment : public OutlookStorage
{
public:
    BYTE* Data;
    ULONG DataLength;
    wstring ContentId;
    wstring Filename;
    int RenderingPosisiton;

    Attachment(IStorage* Storage) : OutlookStorage(Storage)
    {
        this->propHeaderSize = PROPERTIES_STREAM_HEADER_ATTACH_OR_RECIP;
        this->Data = this->GetMapiPropertyBytes(PR_ATTACH_DATA, &(this->DataLength));
        this->ContentId = this->GetMapiPropertyString(PR_ATTACH_CONTENT_ID);
        this->RenderingPosisiton = this->GetMapiPropertyInt32(PR_RENDERING_POSITION);
        wstring filename = this->GetMapiPropertyString(PR_ATTACH_LONG_FILENAME);
        if (filename.empty())
        {
            filename = this->GetMapiPropertyString(PR_ATTACH_FILENAME);
        }
        if (filename.empty())
        {
            filename = this->GetMapiPropertyString(PR_DISPLAY_NAME);
        }
        this->Filename = filename;
    }

    virtual ~Attachment()
    {
        if (this->Data != NULL)
        {
            free(this->Data);
            this->Data = NULL;
        }
    }
};

class Recipient : public OutlookStorage
{
public:
    enum RecipientType
    {
        To,
        CC,
        Unknown
    };
    wstring DisplayName;
    wstring Email;
    string Type;

    Recipient(IStorage* Storage) : OutlookStorage(Storage)
    {
        this->propHeaderSize = PROPERTIES_STREAM_HEADER_ATTACH_OR_RECIP;
        this->DisplayName = this->GetMapiPropertyString(PR_DISPLAY_NAME);
        int recipientType = this->GetMapiPropertyInt32(PR_RECIPIENT_TYPE);
        if (recipientType == MAPI_TO)
        {
            this->Type = "To";
        }
        else if (recipientType == MAPI_CC)
        {
            this->Type = "CC";
        }
        else
        {
            this->Type = "Unknown";
        }
        wstring email = this->GetMapiPropertyString(PR_EMAIL);
        if (email.empty())
        {
            email = this->GetMapiPropertyString(PR_EMAIL_2);
        }
        this->Email = email;
    }

    virtual ~Recipient() { }
};

class Message : public OutlookStorage
{
public:
    vector<Attachment*> Attachments;
    vector<Recipient*> Recipients;
    vector<Message*> Messages;
    wstring From = this->GetMapiPropertyString(PR_SENDER_NAME);
    wstring Subject = this->GetMapiPropertyString(PR_SUBJECT);
    wstring BodyText = this->GetMapiPropertyString(PR_BODY);

    Message(wstring path) : OutlookStorage(path)
    {
        this->LoadStorage(this->Storage);
    }

    Message(IStream* storageStream) : OutlookStorage(storageStream) { }

    Message(IStorage* storage) : OutlookStorage(storage)
    {
        this->propHeaderSize = PROPERTIES_STREAM_HEADER_TOP;
        this->LoadStorage(this->Storage);
    }

    virtual ~Message()
    {
        for (vector<Attachment*>::iterator i = this->Attachments.begin(); i != this->Attachments.end(); i++)
        {
            delete (*i);
        }
        for (vector<Recipient*>::iterator i = this->Recipients.begin(); i != this->Recipients.end(); i++)
        {
            delete (*i);
        }
        for (vector<Message*>::iterator i = this->Messages.begin(); i != this->Messages.end(); i++)
        {
            delete (*i);
        }
        vector<Attachment*>().swap(this->Attachments);
        vector<Recipient*>().swap(this->Recipients);
        vector<Message*>().swap(this->Messages);
    }

    virtual void LoadStorage(IStorage* storage)
    {
        map<LPOLESTR, STATSTG>::iterator i;
        for (i = this->subStorageStatistics.begin(); i != subStorageStatistics.end(); i++)
        {
            IStorage* subStorage;
            wstring v1 = i->first;
            this->Storage->OpenStorage(i->first, NULL, STGM_READ | STGM_SHARE_EXCLUSIVE, NULL, 0, &subStorage);
            if (v1.find(str2wstr(RECIP_STORAGE_PREFIX, CP_UTF8)) == 0)
            {
                Recipient* recipient = new Recipient(subStorage);
                this->Recipients.push_back(recipient);
            }
            else if (v1.find(str2wstr(ATTACH_STORAGE_PREFIX, CP_UTF8)) == 0)
            {
                LoadAttachmentStorage(subStorage);
            }
            else
            {
                subStorage->Release();
            }
        }
    }

    void LoadAttachmentStorage(IStorage* storage)
    {
        Attachment* attachment = new Attachment(storage);
        int attachMethod = attachment->GetMapiPropertyInt32(PR_ATTACH_METHOD);
        if (attachMethod == ATTACH_EMBEDDED_MSG)
        {
            Message* subMsg = new Message(GetMapiPropertyIStorage(PR_ATTACH_DATA));
            subMsg->parentMessage = this;
            subMsg->propHeaderSize = PROPERTIES_STREAM_HEADER_EMBEDED;
            this->Messages.push_back(subMsg);
        }
        else
        {
            this->Attachments.push_back(attachment);
        }
    }
};
