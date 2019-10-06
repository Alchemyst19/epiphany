#!/usr/bin/env python3

'''

Helper for apps that use PDS databases (that were imported into
SQLite3 databases).

Most routines in this module are private to the module (i.e., those
starting with "_").  There's only a handful of public functions.

'''

import re
import pathlib

from datetime import datetime

import PDS

##############################################################################
#
# Public values

# Keys for types of emails
pkey  = 'preferred_emails'
npkey = 'non_preferred_emails'

##############################################################################

# Which database number to use?
# At ECC, the active database is 1.
_database = 1

def _get_db_num():
    return _database

#-----------------------------------------------------------------------------

# These values are not in the database -- they are hard-coded (!)
def _find_member_types():
    member_types = {
        0 : 'Head of Household',
        1 : 'Spouse',
        2 : 'Adult',
        3 : 'Young Adult',
        4 : 'Child',
        5 : 'Other',
    }

    return member_types

#-----------------------------------------------------------------------------

# Normalize some flags to actual Python booleans
def _normalize_boolean(item, src, dest=None) -> None:
    if dest is None:
        dest=src

    if src not in item:
        item[dest] = False
    elif item[src] == '' or item[src] == 0:
        item[dest] = False
        if src != dest:
            del item[src]
    elif item[src] == 1:
        item[dest] = True
        if src != dest:
            del item[src]
    else:
        log.info("Unrecognized {f} value: {v}"
                 .format(f=field, v=item[src]))

#-----------------------------------------------------------------------------

# Represent filenames with a Pathlib object, so that it's equally accessible
# when running on Windows and Linux.
def _normalize_filename(item, src) -> None:
    if src not in item:
        return

    if not item[src]:
        del item[src]
        return

    item[src] = pathlib.PureWindowsPath(item[src])

#-----------------------------------------------------------------------------

def _load_families(pds, columns=None,
                   active_only=True, log=None):
    db_num = _get_db_num()

    if not columns:
        columns = list()
    columns.append('Name')
    columns.append('ParKey')
    columns.append('StreetAddress1')
    columns.append('StreetAddress2')
    columns.append('StreetCityRec')
    columns.append('StreetZip')
    columns.append('PictureFile')
    columns.append('PDSInactive{num}'.format(num=db_num))

    where = ('Fam_DB.CensusFamily{db_num}=1'
             .format(db_num=db_num))
    if active_only:
        where += (' AND '
                  '(Fam_DB.PDSInactive{db_num}=0 OR '
                  'FAM_DB.PDSInactive{db_num} is null)'
                  .format(db_num=db_num))

    families = PDS.read_table(pds, 'Fam_DB', 'FamRecNum',
                              columns=columns, log=log,
                              where=where)

    for f in families.values():
        _normalize_boolean(f, src='PDSInactive{n}'.format(n=db_num),
                    dest="Inactive")
        _normalize_filename(f, src='PictureFile')

    return families

#-----------------------------------------------------------------------------

def _load_members(pds, columns=None,
                  active_only=True, log=None):
    db_num = _get_db_num()

    if not columns:
        columns = list()
    columns.append('Name')
    columns.append('FamRecNum')
    columns.append('DateOfBirth')
    columns.append('MonthOfBirth')
    columns.append('DayOfBirth')
    columns.append('YearOfBirth')
    columns.append('MaritalStatusRec')
    columns.append('MemberType')
    columns.append('PictureFile')
    columns.append('User4DescRec')
    columns.append('Deceased')
    columns.append('PDSInactive{num}'.format(num=db_num))

    where = ('Mem_DB.CensusMember{db_num}=1'
             .format(db_num=db_num))
    if active_only:
        where += (' AND '
                  'Mem_DB.deceased=0 AND '
                  '(Mem_DB.PDSInactive{db_num}=0 OR '
                  'Mem_DB.PDSInactive{db_num} is null)'
                  .format(db_num=db_num))

    members = PDS.read_table(pds, 'Mem_DB', 'MemRecNum',
                             columns=columns, log=log,
                             where=where)

    for m in members.values():
        _normalize_boolean(m, src='Deceased')
        _normalize_boolean(m, src='PDSInactive{n}'.format(n=db_num),
                    dest="Inactive")
        _normalize_filename(m, src='PictureFile')

    return members

#-----------------------------------------------------------------------------

def _link_families_members(families, members):
    # Make a copy because we don't to delete Members from the
    # main/actual list
    members_copy = members.copy()

    for fid, f in families.items():
        family_members = list()
        for mid in members_copy:
            m = members[mid]

            frn = m['FamRecNum']
            if fid == frn:
                family_members.append(m)
                m['family'] = f

        # Delete all the Members we found from the main list of active
        # members (because we already found their families)
        for m in family_members:
            del members_copy[m['MemRecNum']]

        f['members'] = family_members

#-----------------------------------------------------------------------------

def _delete_non_parishioners(families, members):
    to_delete = list()

    # Look for family ParKey >= 10,000
    for fid, f in families.items():
        parkey = int(f['ParKey'])
        if parkey >= 10000:
            f = families[fid]
            for m in f['members']:
                mid = m['MemRecNum']
                del members[mid]

            to_delete.append(fid)

    for fid in to_delete:
        del families[fid]

#-----------------------------------------------------------------------------

def _link_family_emails(families, emails):
    for _, f in families.items():
        f[pkey]  = list()
        f[npkey] = list()

    for _, e in emails.items():
        if not e['FamEmail']:
            continue

        fid = e['MemRecNum']
        if fid not in families:
            continue

        f = families[fid]
        if e['EMailOverMail']:
            key = pkey
        else:
            key = npkey

        addr = '{name} <{addr}>'.format(name=f['Name'],
                                        addr=e['EMailAddress'])
        e['full_address'] = addr
        f[key].append(e)

#-----------------------------------------------------------------------------

def _link_family_city_states(families, city_states):
    for _, f in families.items():
        csid = f['StreetCityRec']
        if csid:
            f['city_state'] = city_states[csid]['CityState']

#-----------------------------------------------------------------------------

def _link_member_types(members, types):
    for _, m in members.items():
        m['type'] = types[m['MemberType']]

#-----------------------------------------------------------------------------

def _link_member_emails(members, emails):
    for _, m in members.items():
        m[pkey]  = list()
        m[npkey] = list()

    for _, e in emails.items():
        if e['FamEmail']:
            continue

        mid = e['MemRecNum']
        if mid not in members:
            continue

        m = members[mid]
        if e['EMailOverMail']:
            key = pkey
        else:
            key = npkey

        addr = '{name} <{addr}>'.format(name=m['email_name'],
                                        addr=e['EMailAddress'])
        e['full_address'] = addr
        m[key].append(e)

#-----------------------------------------------------------------------------

def _link_member_phones(members, phones, phone_types):
    for _, p in phones.items():
        mid = p['Rec']
        if mid not in members:
            continue

        m = members[mid]
        if 'phones' not in m:
            m['phones'] = list()

        ptr = p['PhoneTypeRec']
        phone_type = ''
        if ptr in phone_types:
            phone_type = phone_types[ptr]['Description']

        _normalize_boolean(p, 'Unlisted')
        m['phones'].append({
            'number'   : p['Number'],
            'type'     : phone_type,
            'unlisted' : p['Unlisted'],
        })

#-----------------------------------------------------------------------------

def _link_member_keywords(members, keywords, mem_keywords):
    for _, mk in mem_keywords.items():
        mid = mk['MemRecNum']
        if mid not in members:
            continue

        m = members[mid]
        if 'keywords' not in m:
            m['keywords'] = list()
        keyword = keywords[mk['DescRec']]['Description']
        m['keywords'].append(keyword)

#-----------------------------------------------------------------------------

def _link_member_birth_places(members, birth_places):
    for _, b in birth_places.items():
        mid = b['AskMemNum']
        if mid not in members:
            continue

        m = members[mid]
        m['birth_place'] = b['BirthPlace']

#-----------------------------------------------------------------------------

def _link_member_ministries(members, ministries, mem_ministries, statuses):
    akey = 'active_ministries'
    ikey = 'inactive_ministries'

    for _, m in members.items():
        m[akey] = list()
        m[ikey] = list()

    for _, mm in mem_ministries.items():
        mid = mm['MemRecNum']
        if mid not in members:
            continue
        m = members[mid]

        status_id = mm['StatusDescRec']
        if not status_id:
            continue
        status = statuses[status_id]
        mem_list_name = akey
        if status['Active'] != 1:
            mem_list_name = ikey

        ministry_id = mm['MinDescRec']

        # Deep copy the ministry record so that we can add some more
        # data in it about this specific member
        ministry = ministries[ministry_id].copy()
        ministry['active'] = status['Active']
        ministry['status'] = status['Description']

        m[mem_list_name].append(ministry)

#-----------------------------------------------------------------------------

def _link_member_marital_statuses(members, statuses):
    for _, m in members.items():
        ms = m['MaritalStatusRec']
        if ms:
            m['marital_status'] = statuses[ms]['Description']

#-----------------------------------------------------------------------------

def _link_member_marriage_dates(members, mem_dates, mdtid):
    for _, md in mem_dates.items():
        if md['DescRec'] != mdtid:
            continue

        mid = md['MemRecNum']
        if mid not in members:
            continue
        m = members[mid]
        m['marriage_date'] = md['Date']

#-----------------------------------------------------------------------------

def _link_member_occupations(members, occupations):
    for _, m in members.items():

        oid = m['User4DescRec']
        if not oid:
            continue

        occupation = occupations[oid]
        m['occupation'] = occupation['Description']

#-----------------------------------------------------------------------------

# Transform the list of all family fund history (i.e., individual
# contributions) to be:
#
# families['funding'][fid][year][fund_id], a dictionary containing:
#
# * 'fund': PDS DB entry from FundSetup_DB
# * 'history': array of entries, one per contribution of the family that year
# on that fund, each entry containing a dictionary of:
#     * 'activity': name of fund from FuncAct (don't both copying over
#        other data -- the fund name is really the only important thing)
#     * 'fund_id': same as fund_id index in "funding"
#     * 'year': same as year index in "fundung"
#     * 'item': detailed dictionary of information about the contribution.
#       'FEAmt', 'FEComment', 'FEDate' are probably the only relevant fields
#       from this dictionary.
def _link_family_funds(funds, fund_periods, fund_activities,
                       families, all_family_funds, all_family_fund_rates,
                       all_family_fund_history, log):
    # Make a cross reference dictionary of funds by fund ID+year.  It will be
    # used below.
    fund_xref = dict()
    for period in fund_periods.values():
        fund_id = period['FundNumber']
        fund_year = period['FundYear']
        fund = funds[period['SetupRecNum']]

        if fund_year not in fund_xref:
            fund_xref[fund_year] = dict()
        if fund_id not in fund_xref[fund_year]:
            fund_xref[fund_year][fund_id] = dict()

        fund_xref[fund_year][fund_id] = fund

    # Similarly, make a family fund rate cross reference dictionary indexed by
    # family fund IDs, to be used for direct lookups, below.
    family_fund_rate_xref = dict()
    for family_fund_rate in all_family_fund_rates.values():
        family_fund_id = family_fund_rate['FundRecNum']
        family_fund_rate_xref[family_fund_id] = family_fund_rate

    # Do the main work of this method in a standalone dictionary for simplicity.
    # We'll link it into the main "families" dictionary at the end.
    funding = dict()
    for item in all_family_fund_history.values():
        # Make sure this family is in the families dictionary (e.g., if we only
        # have the active families, make sure this is an active family)
        fid = item['FEFamRec']
        if fid not in families:
            continue

        # Transform the item date string into a datetime.date
        d = datetime.strptime(item['FEDate'], '%Y-%m-%d')
        item['FEDate'] = d

        family_fund = all_family_funds[item['FEFundRec']]
        fund_id     = family_fund['FDFund']
        year        = family_fund['FDYear']
        fund        = fund_xref[year][fund_id]

        # Sometimes activity_id will be None.  Thanks PDS!
        activity_id = item['ActRecNum']
        if activity_id:
            activity = fund_activities[activity_id]['Activity']
        else:
            activity = 'None'

        # If the family pledged, they'll have a fund_rate.  If not, they won't.
        family_fund_id = family_fund['FDRecNum']
        if family_fund_id in family_fund_rate_xref:
            fund_rate = family_fund_rate_xref[family_fund_id]
        else:
            fund_rate = None

        # Create the multi-levels in the output
        if fid not in funding:
            funding[fid] = dict()
        if year not in funding[fid]:
            funding[fid][year] = dict()
        if fund_id not in funding[fid][year]:
            funding[fid][year][fund_id] = {
                "fund"      : fund,
                "fund_rate" : fund_rate,
                "history"   : list(),
            }

        funding[fid][year][fund_id]['history'].append({
            "fund_id"  : fund_id,
            "year"     : year,
            "activity" : activity,
            "item"     : item,
        })

    # Now assign the results back to families[fid]['funding']
    for fid in funding:
        # Make sure this family is in the families dictionary (e.g., if we only
        # have the active families, make sure this is an active family). This is
        # technicaly redundant with above, but hey -- defensive programming,
        # right?
        if fid not in families:
            continue

        families[fid]['funds'] = funding[fid]

#-----------------------------------------------------------------------------

def _find_member_marriage_date_type(date_types):
    for dtid, dt in date_types.items():
        if dt['Description'] == 'Marriage':
            return dtid

    return None

#-----------------------------------------------------------------------------

# A full name will be formatted:
#
#    Last,First{Middle}(Nickname}[Maiden],Prefix,Suffix
#
# If Middle, Nickname, or Maiden are not provided, those terms
# (including "{}", "()", and "[]") are not included.  E.g., if only
# the nickname is provided:
#
#    Squyres,Jeffrey(Jeff)
#
# If Prefix and Suffix are not provided, those terms are not there,
# either (including the commas).  If only Suffix is supplied, then the
# comma will be there for the Prefix, but it will be empty.  Example:
#
#    Squyres,Jeffrey{Michael}(Jeff),,Esq.
#
# There are no cases in Epiphany's database where someone does not
# have both a first and a last name.  So I didn't even bother trying
# to figure out how that would be stored.

def _parse_name(name, log=None):
    parts = name.split(',')
    last = parts[0]

    prefix = None
    if len(parts) > 2:
        prefix = parts[2]
        if prefix == '':
            prefix = None

    suffix = None
    if len(parts) > 3:
        suffix = parts[3]

    # The "more" field may have the middle, nickname, and maiden name.
    # Parse those out.
    first = None
    middle = None
    nickname = None
    maiden = None
    if len(parts) > 1:
        more = parts[1]
        result = re.match('([^\(\{\[]+)', more)
        if result:
            first = result[1]
        else:
            first = 'Unknown'

        result = re.search('\{(.+)\}', more)
        if result:
            middle = result[1]

        result = re.search('\((.+)\)', more)
        if result:
            nickname = result[1]

        result = re.search('\[(.+)\]', more)
        if result:
            maiden = result[1]


    if log:
        log.debug("Last: {l}, First: {f}, Middle: {m}, Nickname: {n}, Maiden: {maiden}, Prefix: {pre}, Suffix: {suff}"
                  .format(l=last,f=first,m=middle,n=nickname,maiden=maiden,pre=prefix,suff=suffix))

    return last, first, middle, nickname, maiden, prefix, suffix

def _parse_member_names(members):
    for _, m in members.items():
        name = m['Name']
        (last, first, middle, nickname, maiden,
         prefix, suffix) = _parse_name(name)

        m['first']    = first
        m['middle']   = middle
        m['last']     = last
        m['nickname'] = nickname
        m['maiden']   = maiden
        m['prefix']   = prefix
        m['suffix']   = suffix

        field = 'full_name'
        m[field]     = ''
        if prefix:
            m[field] += prefix + ' '
        if first:
            m[field] += first + ' '
        if nickname:
            m[field] += '("' + nickname + '") '
        if middle:
            m[field] += middle + ' '
        if last:
            m[field] += last
        if maiden:
            m[field] += ' (maiden: ' + maiden + ')'
        if suffix:
            m[field] += ', ' + suffix

        if nickname:
            m['email_name'] = '{nick} {last}'.format(nick=nickname, last=last)
        else:
            m['email_name'] = '{first} {last}'.format(first=first, last=last)

#-----------------------------------------------------------------------------

def _make_emails_lower_case(emails):
    key = 'EMailAddress'
    for _, e in emails.items():
        addr = e[key].lower()
        e[key] = addr

#-----------------------------------------------------------------------------

# Load PDS Families and Members.  Return them as 2 giant hashes,
# appropriately cross-linked to each other.
def load_families_and_members(filename=None, pds=None,
                              active_only=True, parishioners_only=True,
                              log=None):

    if pds and filename:
        raise Exception("Cannot supply both filename *and* PDS SQLite3 cursor -- only supply one or the other")

    if filename:
        pds = PDS.connect(filename)

    city_states = PDS.read_table(pds, 'City_DB', 'CityRec',
                                 columns=['CityState'], log=log)
    statuses    = PDS.read_table(pds, 'StatusType_DB', 'StatusDescRec',
                                 columns=['Description', 'Active'], log=log)
    ministries  = PDS.read_table(pds, 'MinType_DB', 'MinDescRec',
                                 columns=['Description'], log=log)
    keywords    = PDS.read_table(pds, 'MemKWType_DB', 'DescRec',
                                 columns=['Description'], log=log)
    birth_places= PDS.read_table(pds, 'Ask_DB', 'AskRecNum',
                                 columns=['AskMemNum', 'BirthPlace'], log=log)
    date_places = PDS.read_table(pds, 'DatePlace_DB', 'DatePlaceRecNum',
                                 log=log)
    date_types  = PDS.read_table(pds, 'DateType_DB', 'DescRec',
                                 columns=['Description'], log=log)
    phone_types = PDS.read_table(pds, 'PhoneTyp_DB', 'PhoneTypeRec',
                                 columns=['Description'], log=log)
    emails      = PDS.read_table(pds, 'MemEMail_DB', 'EMailRec',
                                 columns=['MemRecNum', 'EMailAddress',
                                          'EMailOverMail', 'FamEmail'],
                                 log=log)
    mem_phones  = PDS.read_table(pds, 'MemPhone_DB', 'PhoneRec',
                                 columns=['Rec', 'Number', 'PhoneTypeRec', 'Unlisted'],
                                 log=log)
    mem_keywords= PDS.read_table(pds, 'MemKW_DB', 'MemKWRecNum',
                                 columns=['MemRecNum', 'DescRec'],
                                 log=log)
    mem_ministries=PDS.read_table(pds, 'MemMin_DB', 'MemKWRecNum',
                                  columns=['MinDescRec', 'MemRecNum',
                                           'StatusDescRec'],
                                  log=log)
    mem_dates   = PDS.read_table(pds, 'MemDates_DB', 'MemDateRecNum',
                                 columns=['MemRecNum', 'Date',
                                          'DescRec'],
                                 log=log)
    mem_4kw     = PDS.read_table(pds, 'User4KW_DB', 'User4DescRec',
                                 columns=['Description'], log=log)

    relationship_types = PDS.read_table(pds, 'RelType_DB', 'RelDescRec',
                                        columns=['Description'], log=log)
    marital_statuses = PDS.read_table(pds, 'MemStatType_DB', 'MaritalStatusRec',
                                      columns=['Description'], log=log)
    # Descriptions of each fund
    funds = PDS.read_table(pds, 'FundSetup_DB', 'SetupRecNum',
                                      columns=['FundNumber',
                                                'FundKey',
                                                'FundName'], log=log)
    # Each fund also has one or more time periods associated with it
    fund_periods = PDS.read_table(pds, 'FundPeriod_DB', 'FundPeriodRecNum',
                                columns=['SetupRecNum', 'FundNumber',
                                         'FundYear', 'FundStart', 'FundEnd'],
                                log=log)
    # When a Family contributes, each contribution is assocaited with
    # a "funding activity"
    fund_activities = PDS.read_table(pds, 'FundAct_DB', 'ActRecNum',
                                  columns=['FundRecNum',
                                            'GroupName',
                                            'Activity',
                                            'Function',
                                            'GroupOrder',
                                            'pdsorder'], log=log)

    # Families' activities with relation to the established funds (there is one
    # entry for each family for each fund to which that family has contributed).
    fam_funds = PDS.read_table(pds, 'FamFund_DB', 'FDRecNum',
                            columns=['FDFamRec', 'FDYear', 'FDFund',
                                    'FDOrder', 'MemRecNum', 'Comment'],
                            log=log)
    # Pledging information from the family
    fam_fund_rates = PDS.read_table(pds, 'FamFundRate_DB', 'RateRecNum',
                            columns=['FundRecNum', 'FDStartDate', 'FDEndDate',
                                    'FDRate', 'FDRateAdj', 'FDNumber',
                                    'FDPeriod', 'FDTotal',
                                    'Batch', 'BatchDate'])
    # A listing of each individual contribution from each family,
    # cross-referenced to fam_funds.
    fam_fund_history = PDS.read_table(pds, 'FamFundHist_DB', 'FERecNum',
                                columns=['FEDate', 'ActRecNum', 'FEFundRec',
                                        'FEFamRec', 'FEAmt', 'FEBatch',
                                        'MemRecNum', 'FEChk', 'FEComment'],
                                log=log)

    member_types = _find_member_types()
    mdtid        = _find_member_marriage_date_type(date_types)

    _make_emails_lower_case(emails)

    families = _load_families(pds=pds,
                              active_only=active_only,
                              log=log)
    members  = _load_members(pds=pds,
                             active_only=active_only,
                             log=log)

    _link_families_members(families, members)

    if parishioners_only:
        _delete_non_parishioners(families, members)

    _link_family_emails(families, emails)
    _link_family_city_states(families, city_states)

    _parse_member_names(members)
    _link_member_types(members, member_types)
    _link_member_emails(members, emails)
    _link_member_phones(members, mem_phones, phone_types)
    _link_member_keywords(members, keywords, mem_keywords)
    _link_member_birth_places(members, birth_places)
    _link_member_ministries(members, ministries, mem_ministries, statuses)
    _link_member_marital_statuses(members, marital_statuses)
    _link_member_marriage_dates(members, mem_dates, mdtid)
    _link_member_occupations(members, mem_4kw)

    _link_family_funds(funds, fund_periods, fund_activities,
                       families, fam_funds, fam_fund_rates, fam_fund_history,
                       log)

    return pds, families, members

##############################################################################

def _get_sorted_addrs(entries):
    addrs = list()
    for entry in entries:
        addrs.append(entry['EMailAddress'])

    return sorted(addrs)

# If a Member or Family has one or more preferred email addresses,
# return them (as an array).  If there are no preferred email
# addresses, return None.
def find_preferred_email(member_or_family):
    mof = member_or_family
    if pkey in mof and len(mof[pkey]) > 0:
        return _get_sorted_addrs(mof[pkey])
    else:
        return [ ]

# Return either the Member/Family preferred email addresses, or, if
# there are no preferred addresses, return the first (by sorted order)
# non-preferred email address (if it exists).  If no email addresses
# exist, return None.
def find_any_email(member_or_family):
    mof = member_or_family
    addrs = find_preferred_email(mof)
    if addrs:
        return addrs
    elif npkey in mof and len(mof[npkey]) > 0:
        addr = _get_sorted_addrs(mof[npkey])[0]
        return [ addr ]
    else:
        return [ ]
