DEFAULT_SHEET_IDS = {
    "MVA": {
        "2027": "1PyKble2HQUEA3EeL4RDIKdRu_9NkrEwDUh-SDGC70-Y",
    },
    "EH": {
        "2027": "12Fb8oVxTI12tigbl56IV5snoGB8SnFBqlMThNs_vrCo",
    },
}


def sheet_ids_from_environment(environment):
    return {
        "MVA": {
            "2026": environment.get("SHEET_MVA_2026"),
            "2027": environment.get("SHEET_MVA_2027") or DEFAULT_SHEET_IDS["MVA"]["2027"],
        },
        "EH": {
            "2026": environment.get("SHEET_EH_2026"),
            "2027": environment.get("SHEET_EH_2027") or DEFAULT_SHEET_IDS["EH"]["2027"],
        },
    }
