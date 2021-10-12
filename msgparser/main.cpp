#include "common.h"
#include "msgparser.hpp"

BOOL
IsDir(PWCHAR FilePath)
{
    CHAR Temp[MAX_PATH] = { 0 };
    wcstombs_s(NULL, Temp, MAX_PATH, FilePath, MAX_PATH);
    struct stat s;
    if (stat(Temp, &s) == 0)
    {
        if (s.st_mode & S_IFDIR)
        {
            return TRUE;
        }
    }
    return FALSE;
}

void
SaveMessage(Message* outlookMsg)
{
    WCHAR DestPath[MAX_PATH] = { 0 };
    if (IsDir(DEST_DIR))
    {
        RemoveDirectoryW(DEST_DIR);
    }
    SECURITY_ATTRIBUTES SecurityAttributes = { sizeof(SECURITY_ATTRIBUTES), NULL, FALSE };
    CreateDirectoryW(DEST_DIR, &SecurityAttributes);

    for (vector<Attachment*>::iterator i = outlookMsg->Attachments.begin(); i != outlookMsg->Attachments.end(); i++)
    {
        BYTE* attachBytes = (*i)->Data;
        PathCombineW(DestPath, DEST_DIR, (*i)->Filename.c_str());
        fstream attachStream(DestPath, std::fstream::out | std::ios_base::binary);
        attachStream.write((PCHAR)attachBytes, (*i)->DataLength);
        attachStream.close();
    }
    for (vector<Message*>::iterator i = outlookMsg->Messages.begin(); i != outlookMsg->Messages.end(); i++)
    {
        SaveMessage(*i);
    }
}

void
DisplayMessage(Message* outlookMsg)
{
    printf("From: %ws\r\n", outlookMsg->From.c_str());
    printf("Subject: %ws\r\n", outlookMsg->Subject.c_str());
    printf("Body: %ws\r\n", outlookMsg->BodyText.c_str());
    printf("%zd Recipients\r\n", outlookMsg->Recipients.size());
    for (vector<Recipient*>::iterator i = outlookMsg->Recipients.begin(); i != outlookMsg->Recipients.end(); i++)
    {
        printf("%s:%ws\r\n", (*i)->Type.c_str(), (*i)->Email.c_str());
    }
    printf("%zd Attachments\r\n", outlookMsg->Attachments.size());
    for (vector<Attachment*>::iterator i = outlookMsg->Attachments.begin(); i != outlookMsg->Attachments.end(); i++)
    {
        printf("%ws          \r\n", (*i)->Filename.c_str());
    }
    printf("%zd Messages\r\n", outlookMsg->Messages.size());
    for (vector<Message*>::iterator i = outlookMsg->Messages.begin(); i != outlookMsg->Messages.end(); i++)
    {
        DisplayMessage(*i);
    }
}

void
main(int argc, char* argv[])
{
    setlocale(LC_ALL, "Chinese");

    WCHAR DestPath[MAX_PATH] = { 0 };
    Message* outlookMsg = new Message(FILE_PATH);
    DisplayMessage(outlookMsg);
    SaveMessage(outlookMsg);
    delete outlookMsg;

    getchar();
}
