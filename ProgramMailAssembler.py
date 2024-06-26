from __future__ import annotations
from typing import Optional

from datetime import datetime
import re
import os

from HelpersPackage import FindBracketedText, MessageLog, ReadListAsParmDict, ParmDict, GetParmFromParmDict
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
    mailFormat=GetParmFromParmDict(parameters,"MailFormat").lower().strip()

    # Open the schedule markup file
    reportsdir=GetParmFromParmDict(parameters,"ProgramAnalyzerReportsdir")
    schedPath=OpenProgramFile(f"Program participant schedules.xml", reportsdir)
    Log(f'OpenProgramFile("Program participant schedules.xml", {reportsdir}, ".") yielded {schedPath}')   # Concatenated strings...
    if not schedPath:
        LogError(f'OpenProgramFile of {schedPath} failed')
        exit(999)
    with open(schedPath, "r") as file:
        markuplines=file.read()
    # Remove newlines *outside* markup
    markuplines=markuplines.replace(">\n<", "><")

    if not CheckBalance(markuplines):
        Log(f'CheckBalance failed')
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
    mainNode=Node("Main", markuplines)
    mainNode.Resolve()

    # Now read the People table
    # Format: <person>pppp</person> (repeated, one line per person)
    # pppp: <header>value</header>  (repeated, one for each column in the people tab)


    ppPath=OpenProgramFile("Program participants.xml", reportsdir)
    Log(f'OpenProgramFile("Program participant schedules.xml", "{reportsdir}", ".") yielded {ppPath}')
    if not ppPath:
        MessageLog(f'OpenProgramFile of {ppPath} failed')
        exit(999)
    with open(ppPath, "r") as file:
        peoplefile=file.read()
    peoplelines: list[str]=[]
    while len(peoplefile) > 0:
        _, tag, line, peoplefile=FindAnyBracketedText(peoplefile)
        peoplelines.append(line)

    # A dictionary of people, keyed by the person's full name
    # Each person's value is a dictionary of column values from the people tab
    people=ParmDict(CaseInsensitiveCompare=True, IgnoreSpacesCompare=True)
    for line in peoplelines:
        d=ParmDict(CaseInsensitiveCompare=True, IgnoreSpacesCompare=True)
        while len(line) > 0:
            _, header, value, line=FindAnyBracketedText(line)
            Log(f"{header=}  {value=}")
            d[header]=value
        if len(d) > 0:
            if d.Exists("full name"):
                people[d["full name"]]=d
            else:
                LogError(f"While reading 'Program participants.xml', unable to find a Full Name for \n{line}\n")
                LogError(f"ParmDict={[x for x in d]}\n\n")

    # Read the email template.  It consists of two XMLish items, the selection criterion and the email body
    # Things in [[double brackets]] will be replaced by the corresponding cell from the person's row People page or, in the case of [[schedule]],
    # with the person's schedule.
    templatePath=OpenProgramFile(GetParmFromParmDict(parameters, "PMATemplateFile", "."), '.')
    Log(f"Template is {templatePath}")
    if templatePath is None:
        MessageLog(f"Template file {templatePath} could not be opened")
        exit(999)
    with open(templatePath, "r", encoding="UTF-8") as file:
        template=file.read()

    if not CheckBalance(template):
        MessageLog("The template failed the CheckBalance() test -- it seems to have unbalanced HTML")
        return

    # Read the selection criterion
    # Note that the selection's header value may be empty, but it must be present, as must a (possibly empty) value
    selection, template=FindBracketedText(template, "select", stripHtml=False)
    if len(selection) == 0:
        MessageLog(f"Template does not contain a <selection>...</selection> element")
        return
    header, selection=FindBracketedText(selection, "header", stripHtml=False)
    header=header.strip().lower()
    if len(header) == 0:
        MessageLog(f"<select> element does not contain a <header>...</header> element>")
        return
    selectionvalue, selection=FindBracketedText(selection, "value", stripHtml=False)
    selectionvalue=selectionvalue.strip()
    if len(selectionvalue) == 0:
        MessageLog(f"<selection> element does not contain a <value>...</value> element>")
        return

    # Read the email body
    emailbody, template=FindBracketedText(template, "email body", stripHtml=False)
    if len(emailbody) == 0:
        MessageLog(f"Template does not contain an <email body>...</email body> element>")
        return

    # OK, time to produce the output
    # We loop through all the people who have schedules, and generate emails for those who match the selection criterion.
    # The email file is also XMLish:
    # <person>
    # <email>email address</email>
    # <contents>letter...<contents>
    # </person>  ...and repeated


    # Read the name of the input file to be used
    inputFileName, template=FindBracketedText(template, "inputFileName", stripHtml=False)
    if len(inputFileName) == 0:
        inputFileName="Program participant schedules email.txt"

    with open(inputFileName, "w") as file:
        print(f"# {datetime.now()}\n", file=file)
        for person in mainNode:
            fullname=person["full name"]
            if not people.Exists(fullname):
                LogError(f"For {fullname}, {person['full name']=} not in People -- skipped.")
                continue

            peopledata=people[fullname]
            if not peopledata.Exists(header):
                LogError(f"For {fullname}, {header=} not in People's column headers -- skipped.")
                continue
            headervalue=peopledata[header]
            if headervalue.strip().lower() != selectionvalue.lower():
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
                    # The tag [[schedule]] is special and fetched from a bunch of keys in the person's schedule structure
                    # The others
                    if tag == "schedule":
                        items=""
                        for attribute in person.List:

                            if attribute.Key == "item":
                                title=""
                                participants=""
                                precis=""
                                equipment=""
                                for subatt in attribute.List:
                                    match subatt.Key:
                                        case "title":
                                            title=subatt.Text
                                        case "participants":
                                            participants=subatt.Text
                                        case "precis":
                                            precis=subatt.Text
                                        case "equipment":
                                            equipment=subatt.Text
                                # Now format this item for the email
                                if mailFormat == "html":
                                    item=f"<p><b>{title}</b></p>\n<p>{participants}</p>\n"
                                    if len(equipment) > 0:
                                        item+=f"<p>equipment: {equipment}</p>\n"
                                    if len(precis) > 0:
                                        item+=f"<p>{precis}</p>\n"
                                else:
                                    item=f"{title}\n{participants}\n"
                                    if len(equipment) > 0:
                                        item+=f"equipment: {equipment}\n"
                                    if len(precis) > 0:
                                        item+=f"{precis}\n"
                                # Add the item to the items text block
                                items=items+item
                                if mailFormat == "html":
                                    items=items+"<p>"
                                items=items+"\n"
                                continue
                        # Assemble the opening material, the items text block ansd the closing aterian into a message
                        thismail=start+items+trail

                    # All other tags come from columns of the people tab
                    else:
                        # If the tag is of the form xxx|yyy|xxx, we pass the prefix and suffix through if the center part is non-empty
                        prefix=suffix=""
                        if tag.count("|") == 2:
                            prefix, tag, suffix=tag.split("|")

                        val=people[fullname][tag]
                        if val is None:
                            MessageLog(f"Can't find {tag=} in people.keys() for {fullname}\nAborting execution.")
                            return
                        if len(val)> 0:
                            val=prefix+val+suffix
                        thismail=start+val+trail

            file.write(thismail+"\n")

            file.write(f"</content>")
            file.write(f"</email-message>\n\n\n")



    LogDisplayErrorsIfAny()


class Node:
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
    #Log(f"\nCheckBalance:  {s=}")

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
            m=re.match("^([a-zA-Z0-9])\\s", delim)   # Check for cases like <a http=...> -- the delim is just the a
            if m is not None:
                delim=m.groups()[0]
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
    m1=re.match("^[^<]*?<([^<>\\[\\]]*?)>", s)
    m2=re.match("^[^\[]*?\[\[([^<>\\[\\]]]*?)]", s)

    # Neither found means we're done.
    if m1 is None and m2 is None:
        #Log(f"LocateNextDelimiter: m1=m2=None")
        return "", ""

    if m1 is not None and m2 is None:
        #Log(f"LocateNextDelimiter: m1 ends at {m1.regs[0][1]}")
        return m1.groups()[0], s[m1.regs[0][1]:]

    if m1 is None and m2 is not None:
        #Log(f"LocateNextDelimiter: m2 ends at {m2.regs[0][1]}")
        return "[[", s[m2.regs[0][1]:]

    # Both found. Which is first?
    #Log(f"LocateNextDelimiter: m1 ends at {m1.regs[0][1]} and m2 ends at {m2.regs[0][1]}")
    if m1.regs[0][1] < m2.regs[0][1]:
        return m1.groups()[0], s[m1.regs[0][1]:]
    else:
        return "[[", s[m2.regs[0][1]:]

#-------------------------------------------------
# Search for a Program file and return its path.
# Look first in the location specified by path.  Failing that, look in defaultDir.  Failing that look in the CWD.
def OpenProgramFile(fname: str, path: str, report=True) -> Optional[str]:
    if fname is None:
        MessageLog(f"OpenProgramFile: fname is None, {path=}")
        return None

    if path is not None:
        pathname=os.path.join(path, fname)
        if os.path.exists(pathname):
            return pathname

    pathname=os.path.join(".", fname)
    if os.path.exists(pathname):
        return pathname

    if os.path.exists(fname):
        return fname

    if report:
        if path != ".":
            MessageLog(f"Can't find '{fname}': checked '{path}' and './'")
        else:
            MessageLog(f"Can't find '{fname}': checked './'")

    return None

#=====================================================================================
# Note this is a carient of a method from HelpersFile
# Find first text bracketed by <anything>...</anything>
# Return a tuple consisting of:
#   Any leading material
#   The name of the first pair of brackets found
#   The contents of the first pair of brackets found
#   The remainder of the input string
# Note that this is a *non-greedy* scanner
# Note also that it is not very tolerant of errors in the bracketing, just dropping things on the floor
def FindAnyBracketedText(s: str) -> tuple[str, str, str, str]:

    pattern=r"^(.*?)<([a-zA-Z0-9 ]+)[^>]*?>(.*?)<\/\2>"
    m=re.search(pattern, s,  re.DOTALL)
    if m is None:
        return s, "", "", ""

    x=m.group(1), m.group(2), m.group(3), s[m.regs[0][1]:]
    return x


######################################
# Run main()
#
if __name__ == "__main__":
    main()