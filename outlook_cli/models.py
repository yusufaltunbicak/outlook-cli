from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime


@dataclass
class EmailAddress:
    name: str
    address: str

    @classmethod
    def from_api(cls, data: dict) -> EmailAddress:
        ea = data.get("EmailAddress", data)
        return cls(name=ea.get("Name", ""), address=ea.get("Address", ""))

    def __str__(self) -> str:
        if self.name:
            return f"{self.name} <{self.address}>"
        return self.address


@dataclass
class Email:
    id: str
    subject: str
    sender: EmailAddress
    to: list[EmailAddress]
    cc: list[EmailAddress]
    received: datetime
    preview: str
    body: str
    body_type: str  # "Text" or "HTML"
    is_read: bool
    has_attachments: bool
    importance: str
    conversation_id: str
    categories: list[str] = field(default_factory=list)
    scheduled_send: datetime | None = None
    # display number assigned by CLI
    display_num: int = 0

    @classmethod
    def from_api(cls, data: dict) -> Email:
        scheduled_send = None
        for prop in data.get("SingleValueExtendedProperties", []):
            if "0x3FEF" in prop.get("PropertyId", ""):
                scheduled_send = _parse_dt(prop.get("Value", ""))
                break

        return cls(
            id=data["Id"],
            subject=data.get("Subject", "(No Subject)"),
            sender=EmailAddress.from_api(data.get("From", {}).get("EmailAddress", {"Name": "", "Address": ""})),
            to=[EmailAddress.from_api(r) for r in data.get("ToRecipients", [])],
            cc=[EmailAddress.from_api(r) for r in data.get("CcRecipients", [])],
            received=_parse_dt(data.get("ReceivedDateTime", "")),
            preview=data.get("BodyPreview", ""),
            body=data.get("Body", {}).get("Content", ""),
            body_type=data.get("Body", {}).get("ContentType", "Text"),
            is_read=data.get("IsRead", False),
            has_attachments=data.get("HasAttachments", False),
            importance=data.get("Importance", "Normal"),
            conversation_id=data.get("ConversationId", ""),
            categories=data.get("Categories", []),
            scheduled_send=scheduled_send,
        )


@dataclass
class Folder:
    id: str
    name: str
    unread_count: int
    total_count: int
    parent_folder_id: str

    @classmethod
    def from_api(cls, data: dict) -> Folder:
        return cls(
            id=data["Id"],
            name=data.get("DisplayName", ""),
            unread_count=data.get("UnreadItemCount", 0),
            total_count=data.get("TotalItemCount", 0),
            parent_folder_id=data.get("ParentFolderId", ""),
        )


@dataclass
class Attachment:
    id: str
    name: str
    content_type: str
    size: int
    is_inline: bool
    content_bytes: str | None = None  # base64

    @classmethod
    def from_api(cls, data: dict) -> Attachment:
        return cls(
            id=data["Id"],
            name=data.get("Name", ""),
            content_type=data.get("ContentType", ""),
            size=data.get("Size", 0),
            is_inline=data.get("IsInline", False),
            content_bytes=data.get("ContentBytes"),
        )


@dataclass
class Attendee:
    email: EmailAddress
    type: str  # "Required", "Optional", "Resource"
    response: str  # "None", "Accepted", "Declined", "TentativelyAccepted", "NotResponded"

    @classmethod
    def from_api(cls, data: dict) -> Attendee:
        return cls(
            email=EmailAddress.from_api(data.get("EmailAddress", {})),
            type=data.get("Type", "Required"),
            response=data.get("Status", {}).get("Response", "None"),
        )


@dataclass
class Event:
    id: str
    subject: str
    start: datetime
    end: datetime
    location: str
    organizer: EmailAddress
    is_all_day: bool
    body_preview: str
    body: str
    body_type: str
    attendees: list[Attendee] = field(default_factory=list)
    categories: list[str] = field(default_factory=list)
    show_as: str = "Busy"
    sensitivity: str = "Normal"
    is_cancelled: bool = False
    response_status: str = ""
    web_link: str = ""
    is_online_meeting: bool = False
    online_meeting_url: str = ""
    recurrence: dict | None = None
    event_type: str = ""  # "SingleInstance", "Occurrence", "Exception", "SeriesMaster"
    series_master_id: str = ""
    display_num: int = 0

    @classmethod
    def from_api(cls, data: dict) -> Event:
        org_data = data.get("Organizer", {}).get("EmailAddress", {"Name": "", "Address": ""})
        attendees = [Attendee.from_api(a) for a in data.get("Attendees", [])]
        online_url = ""
        if data.get("OnlineMeeting"):
            online_url = data["OnlineMeeting"].get("JoinUrl", "")
        return cls(
            id=data["Id"],
            subject=data.get("Subject", "(No Subject)"),
            start=_parse_dt(data.get("Start", {}).get("DateTime", "")),
            end=_parse_dt(data.get("End", {}).get("DateTime", "")),
            location=data.get("Location", {}).get("DisplayName", ""),
            organizer=EmailAddress.from_api(org_data),
            is_all_day=data.get("IsAllDay", False),
            body_preview=data.get("BodyPreview", ""),
            body=data.get("Body", {}).get("Content", ""),
            body_type=data.get("Body", {}).get("ContentType", "Text"),
            attendees=attendees,
            categories=data.get("Categories", []),
            show_as=data.get("ShowAs", "Busy"),
            sensitivity=data.get("Sensitivity", "Normal"),
            is_cancelled=data.get("IsCancelled", False),
            response_status=data.get("ResponseStatus", {}).get("Response", ""),
            web_link=data.get("WebLink", ""),
            is_online_meeting=data.get("IsOnlineMeeting", False),
            online_meeting_url=online_url,
            recurrence=data.get("Recurrence"),
            event_type=data.get("Type", ""),
            series_master_id=data.get("SeriesMasterId", ""),
        )


@dataclass
class Contact:
    id: str
    display_name: str
    given_name: str
    surname: str
    email_addresses: list[EmailAddress]
    company: str
    job_title: str

    @classmethod
    def from_api(cls, data: dict) -> Contact:
        emails = [
            EmailAddress(name=e.get("Name", ""), address=e.get("Address", ""))
            for e in data.get("EmailAddresses", [])
        ]
        return cls(
            id=data["Id"],
            display_name=data.get("DisplayName", ""),
            given_name=data.get("GivenName", ""),
            surname=data.get("Surname", ""),
            email_addresses=emails,
            company=data.get("CompanyName", ""),
            job_title=data.get("JobTitle", ""),
        )


def _parse_dt(s: str) -> datetime:
    if not s:
        return datetime.min
    # Outlook returns ISO 8601 with or without timezone
    s = s.replace("Z", "+00:00")
    try:
        return datetime.fromisoformat(s)
    except ValueError:
        return datetime.min
