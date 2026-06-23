"""Shared constants for project name extraction."""

# Garbage words/patterns that indicate a label, not a project name
GARBAGE_WORDS = {
    "title", "projectsystem", "projsys", "pbswbs", "lrusystem", "systempartno",
    "criticallevel", "documentclassification", "copyno", "noofpage",
    "designagencylogo", "namedesignation", "signature", "preparedby",
    "reviewedby", "approvedby", "issueno", "revno", "issuedate", "docno",
    "page", "doc", "document", "file", "issue", "rev", "revision", "date",
    "name", "designation", "agency", "restricted", "secret", "logo", "copy",
    "number", "section", "reference", "subject", "project",
    # additional low-signal words
    "matter", "report", "draft", "note", "notes", "unknown", "untitled",
    "na", "none", "null", "tbd", "tbc", "sheet", "form", "ref", "client",
    "version", "description", "details", "information",
}

# Project name length & complexity constraints
PROJECT_NAME_MIN_LENGTH = 5
PROJECT_NAME_MAX_LENGTH = 150
PROJECT_NAME_MIN_WORDS = 2
