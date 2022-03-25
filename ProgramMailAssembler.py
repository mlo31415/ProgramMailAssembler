from __future__ import annotations
from typing import Optional

from datetime import datetime

from HelpersPackage import FindAnyBracketedText


def main():
    # Open the schedule markup file
    markuplines=""
    with open("Program participant schedules.xml", "r") as file:
        markuplines=file.read()
    # Remove newlines *outside* markup
    markuplines=markuplines.replace(">\n<", "><")

    # <person>xxxx</person>
    # <person>xxxx</person>
    # <person>xxxx</person>

    # Each xxxx is:
    # <fullname>fffff</fullname>
    # <email>eeeee</email>
    # <item>iiii</item>
    # <item>iiii</item>
    # <item>iiii</item>
    # <item>iiii</item>

    # Each item is:
    # <title>ttt</title>
    # <participants>pppp</participants>
    # <precis>yyyyy</precis>

    # Markup is a dict keyed by the <name></name> with contents the contained markup and rooted at "main"
    main=Node("Main", markuplines).Resolve()

    # Now we have the whole schmeer in a hierarchical structure
    # Generate the output
    with open("Program participant schedules email.txt", "w") as file:
        print(f"# {datetime.now()}\n", file=file)
        for person in main:
            fullname=""
            emailAddr=""
            items=""
            for attribute in person.List:
                if attribute.Key == "email":
                    emailAddr=attribute.Text
                    continue
                if attribute.Key == "fullname":
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
            print(f"<email>", file=file)
            print(f"<email-address>{emailAddr}</email-address>", file=file)
            print(f"<content>", file=file)
            print(f"Dear {fullname}\nHere's yer schedule:\n{items}", file=file)
            print(f"</content>", file=file)
            print(f"</email>\n\n", file=file)




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

    def __getitem__(self, i: int):
        return self._value[i]

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
                    print(f"[({key}, {trail})]")
                    self._value=trail
                    return self
                if trail == "":
                    break
            out.append(Node(bracket, contents).Resolve())
            text=trail

        print(f"[({key}, {len(out)=})]")
        self._value=out
        return self


if __name__ == "__main__":
    main()