from __future__ import annotations
from typing import Optional

from datetime import datetime
import re
import os

from HelpersPackage import FindAnyBracketedText, MessageLog, ReadListAsParmDict, ParmDict
from Log import Log, LogError, LogDisplayErrorsIfAny


#******************************************************************************************************************************************************
#
#       ProgramMailAssembler
#
# PMA takes a pair of XML files created by ProgramAnalyzer and a template file created by the user and creates an XML file for input to ProgramMailer.
# This output file contains the emails to be sent in a very simple structure.  The user can edit this file before submitting it to PM to be mailed.
#
#******************************************************************************************************************************************************

def main():

    parameters=ReadListAsParmDict('parameters.txt', isFatal=True)
    if parameters is None or len(parameters) == 0:
        MessageLog(f"Can't open/read {os.getcwd()}/parameters.txt")
        exit(999)

    # Open the schedule markup file
    schedPath=OpenProgramFile("Program participant schedules.xml", parameters["ProgramAnalyzerReportsdir"], ".")
    if not schedPath:
        exit(999)
    with open(schedPath, "r") as file:
        markuplines=file.read()
    # Remove newlines *outside* markup
    markuplines=markuplines.replace(">\n<", "><")

    if not CheckBalance(markuplines):
        return

    # <person>xxxx</person>
    # ...
    # <person>xxxx</person>

    # Each xxxx is:
    # <fullname>fffff</fullname>
    # <email>eeeee</email>
    # <item>iiii</item>
    # ...
    # <item>iiii</item>

    # Each iiii is:
    # <title>ttt</title>
    # <participants>pppp</participants>
    # <precis>yyyyy</precis>

    # Markup is a dict keyed by the <name></name> with contents the contained markup and rooted at "main"
    main=Node("Main", markuplines)
    main.Resolve()

    # Now read the People table
    # Format: <person>pppp</person> (repeated)
    # pppp: <header>value</header> repeated for each column

    ppPath=OpenProgramFile("Program participants.xml", parameters["ProgramAnalyzerReportsdir"], ".")
    if not ppPath:
        exit(999)
    with open(ppPath, "r") as file:
        peoplefile=file.read()
    peoplelines: list[str]=[]
    while len(peoplefile) > 0:
        _, tag, line, peoplefile=FindAnyBracketedText(peoplefile)
        peoplelines.append(line)

    # A dictionary of people
    # Each person's value is a dictionary of column values
    people=ParmDict(CaseInsensitiveCompare=True, IgnoreSpacesCompare=True)
    for line in peoplelines:
        d=ParmDict(CaseInsensitiveCompare=True, IgnoreSpacesCompare=True)
        while len(line) > 0:
            _, header, value, line=FindAnyBracketedText(line)
            #Log(f"{header=}  {value=}")
            d[header.lower()]=value
        if d.Exists("full name"):
            people[d["full name"]]=d

    # Read the email template.  It consists of two XMLish items, the selection criterion and the email body
    # Things in [[double brackets]] will be replaced by the corresponding cell from the person's row People page or, in the case of [[schedule]],
    # with the person's schedule.
    templatePath=OpenProgramFile(parameters["PMATemplateFile"], ".", ".")
    if templatePath is None:
        exit(999)
    with open("Template.xml", "r") as file:
        template=file.read()

    if not CheckBalance(template):
        return

    # Read the selection criterion
    # Note that the selection's header value may be empty, but it must be present, as must a (possibly empty) value
    _, tag, selection, template=FindAnyBracketedText(template)
    if tag != "select":
        MessageLog(f"First item in template is not the selection: {tag=}  and {selection=}")
        return
    _, tag, header, selection=FindAnyBracketedText(selection)
    header=header.strip().lower()
    if tag != "header":
        MessageLog(f"First item in select specification is not the header: {tag=}  and {selection=}")
        return
    _, tag, selectionvalue, selection=FindAnyBracketedText(selection)
    selectionvalue=selectionvalue.strip()
    if tag != "value":
        MessageLog(f"Second item in select specification is not the selection value: {tag=}  and {selection=}")
        return

    # Read the email body
    _, tag, emailbody, template=FindAnyBracketedText(template)
    if tag != "email body":
        MessageLog(f"Second item in template is not the email body: {tag=}  and {emailbody=}")
        return

    # OK, time to produce the output
    # We loop through all the people who have schedules, and generate emails for those who match the selection criterion.
    # The email file is also XMLish:
    # <person>
    # <email>email address</email>
    # <contents>letter...<contents>
    # </person>  ...and repeated

    with open("Program participant schedules email.txt", "w") as file:
        print(f"# {datetime.now()}\n", file=file)
        for person in main:
            fullname=person["full name"]
            if not people.Exists(fullname):
                LogError(f"For {fullname}, {person['full name']=} not in People -- skipped.")
                continue

            peopledata=people[fullname]
            if not peopledata.Exists(header):
                LogError(f"For {fullname}, {header=} not in People's column headers -- skipped.")
                continue
            headervalue=peopledata[header]
            if headervalue.strip() != selectionvalue:
                Log(f"For {fullname}, {headervalue=} does not match {selectionvalue=} -- skipped.")
                continue

            file.write(f"<email-message>")
            emailAddr=person["email"]
            file.write(f"<email-address>{emailAddr}</email-address>")
            file.write(f"<content>")

            # Now substitute into the email body from the template and write it
            # We scan for all [[xxx]] and replace it with people[person][xxx]
            thismail=emailbody
            while thismail.find("[[") >= 0:
                loc1=thismail.find("[[")
                loc2=thismail[loc1:].find("]]")
                if loc1 >= 0 and loc2 >= 0:
                    start=thismail[:loc1]
                    tag=thismail[loc1+2:loc1+loc2].lower()
                    trail=thismail[loc1+loc2+2:]
                    # Substitute content for the tag
                    if tag == "schedule":
                        items=""
                        for attribute in person.List:
                            if attribute.Key == "full name":
                                fullname=attribute.Text
                                continue
                            if attribute.Key == "item":
                                title=""
                                participants=""
                                precis=""
                                for subatt in attribute.List:
                                    if subatt.Key == "title":
                                        title=subatt.Text
                                    if subatt.Key == "participants":
                                        participants=subatt.Text
                                    if subatt.Key == "precis":
                                        precis=subatt.Text
                                item=f"{title}\n{participants}\n"
                                if len(precis) > 0:
                                    item+=f"{precis}\n"
                                items=items+item+"\n"
                                continue
                        thismail=start+items+trail
                    else:   # All other tags come from the people tab
                        if tag not in next(iter(people.values())):  # Kludge to get the keys of the inner dictionary
                            Log(f"Can't find {tag=} in people.keys()", isError=True)
                            break
                        thismail=start+people[person['full name']][tag]+trail

            file.write(thismail+"\n")

            file.write(f"</content>")
            file.write(f"</email-message>\n\n\n")

    LogDisplayErrorsIfAny()




class Node():
    def __init__(self, key: str, value: str|list[Node] = ""):
        self._key=key

        if type(value) == str:
            self._value=value
            return
        if type(value) == list:
            self._value=value
            return

        assert False

    def __len__(self) -> int:
        assert type(self._value) != Node
        return len(self._value)

    def __getitem__(self, index):
        if type(index) is int:
            return self._value[index]
        if type(index) is str:
            for node in self.List:
                if node.Key == index.lower():
                    return node.Text
        assert False



    @property
    def IsText(self) -> bool:
        return type(self._value) == str

    @property
    def Key(self) -> str:
        return self._key

    @property
    def List(self) -> list[Node]:
        if type(self._value) == list:
            return self._value
        return []

    @property
    def Text(self) -> str:
        if type(self._value) == str:
            return self._value
        return ""

    # Recursively parse the markup
    def Resolve(self) -> Node:
        # Replace the list of strings with a list of dicts for each xxx and then call resolve for each of those
        # Find the first (perhaps only) markup in the list of strings

        key=self._key
        text=self._value

        out: list[Node]=[]
        while len(text) > 0:
            lead, bracket, contents, trail=FindAnyBracketedText(text)
            if bracket == "" and contents == "":
                if trail != "":
                    #print(f"[({key}, {trail})]")
                    self._value=trail
                    return self
                if trail == "":
                    break
            node=Node(bracket, contents)
            node.Resolve()
            out.append(node)
            text=trail

        #print(f"[({key}, {len(out)=})]")
        if out:
            self._value=out
        return self

#-------------------------------------------
# Check a string to make sure that it has balanced and properly nested <xxx></xxx> and [[]]s
# Log errors
def CheckBalance(s: str) -> bool:
    Log(f"\nCheckBalance:  {s=}")

    nesting: list[str]=[]

    # Remove the line ends as they just make Regex harder.
    s=s.replace("\n", " ")

    while s:
        delim, s=LocateNextDelimiter(s)
        #Log(f"CheckBalance:  {delim=}    {s=}")
        if (delim is None or delim == "") and nesting:
            MessageLog(f"Template error: Unbalanced delimiters found around '{s}\nProgramMailAssembler terminated.")
            return False

        if delim == "":
            return True

        # Is this a new opening delimiter?
        if delim == "[[" or delim[0] != "/":
            nesting.append(delim)
            #Log(f"CheckBalance: push '{delim}'   {s=}")
            continue

        # We have a delimiter and it is not an opening delim, so it much be a closing delim.  Is there anything left on the stack to match?
        if not nesting:
            MessageLog(f"CheckBalance: Error -- missing ]] near '{s}\nProgramMailAssembler terminated.")
            return False

        top=nesting.pop()
        #Log(f"CheckBalance: pop '{top}'   {s=}")

        if delim == "]]":
            if top != "[[":
                MessageLog(f"CheckBalance: Error -- Unbalanced [[]] near '{s}\nProgramMailAssembler terminated.")
                return False
            continue

        if delim[0] == "/":
            if top == delim[1:]:
                continue
            MessageLog(f"CheckBalance: Error -- Unbalanced <>...</> near '{s}\nProgramMailAssembler terminated.")
            return False

    return True

# Scan for the next opening or closing delimiter.
# Return a tuple of the delimiter found and the remaining text
# For <xxx>, the delimiter returned is xxx. For [[xxx]], the delimiter returned is [[
# Return (None, str) on error
def LocateNextDelimiter(s: str) -> tuple[Optional[str], str]:
    if not s:
        return "", ""

    # Match <stuff> followed by "<" followed by stuff not containing delimiters, followed by ">", followed by stuff
    m1=re.match("^[^<]*?<([^<>\[\]]*?)>", s)
    m2=re.match("^[^\[]*?\[\[([^<>\[\]]]*?)]", s)

    # Neither found means we're done.
    if m1 is None and m2 is None:
        #Log(f"LocateNextDelimiter: m1=m2=None")
        return "", ""

    if m1 is not None and m2 is None:
        i=0
        #Log(f"LocateNextDelimiter: m1 ends at {m1.regs[0][1]}")
        return m1.groups()[0], s[m1.regs[0][1]:]

    if m1 is None and m2 is not None:
        i=0
        #Log(f"LocateNextDelimiter: m2 ends at {m2.regs[0][1]}")
        return "[[", s[m2.regs[0][1]:]

    # Both found. Which is first?
    #Log(f"LocateNextDelimiter: m1 ends at {m1.regs[0][1]} and m2 ends at {m2.regs[0][1]}")
    if m1.regs[0][1] < m2.regs[0][1]:
        return m1.groups()[0], s[m1.regs[0][1]:]
    else:
        return "[[", s[m2.regs[0][1]:]

# Search for a Program file and return its path.
# Look first in the location specified by path.  Failing that, look in defaultDir.  Failing that look in the CWD.
def OpenProgramFile(fname: str, path: str, defaultDir: str, report=True) -> Optional[str]:
    if fname is None:
        MessageLog(f"OpenProgramFile: fname is None, {path=}")
        return None

    if path is not None:
        pathname=os.path.join(path, fname)
        if os.path.exists(pathname):
            return pathname

    pathname=os.path.join(defaultDir, fname)
    if os.path.exists(pathname):
        return pathname

    if os.path.exists(fname):
        return fname

    if report:
        if defaultDir != "." and path != ".":
            MessageLog(f"Can't find '{fname}': checked '{path}', '{defaultDir}' and './'")
        elif path != ".":
            MessageLog(f"Can't find '{fname}': checked '{path}' and './'")
        else:
            MessageLog(f"Can't find '{fname}': checked './'")

    return None

######################################
# Run main()
#
if __name__ == "__main__":
    main()