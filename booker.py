#!/usr/bin/env python3

# Copyright (c) 2017, Bartlomiej Puget <larhard@gmail.com>
# All rights reserved.
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
#     * Redistributions of source code must retain the above copyright notice,
#       this list of conditions and the following disclaimer.
#
#     * Redistributions in binary form must reproduce the above copyright notice,
#       this list of conditions and the following disclaimer in the documentation
#       and/or other materials provided with the distribution.
#
#     * Neither the name of the Bartlomiej Puget nor the names of its
#       contributors may be used to endorse or promote products derived from this
#       software without specific prior written permission.
#
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
# ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
# WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL BARTLOMIEJ PUGET BE LIABLE FOR ANY DIRECT,
# INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING,
# BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
# DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF
# LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE
# OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF
# ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.


import argparse
import collections
import datetime
import logging
import os
import re
import subprocess
import sys
import tempfile

log = logging.getLogger("Booker")


class Paper:
    def __init__(self, size):
        self.size = size.lower()

    @property
    def upper(self):
        return self.size.upper()

    @property
    def lower(self):
        return self.size

    @property
    def latex(self):
        return "{}paper".format(self.size)


def gen_raw_selection(selection):
    if selection is None:
        return None

    return ",".join(str(k) for k in selection)


def mktemp(prefix="booker-", suffix="", dir=None, text=False):
    fd, path = tempfile.mkstemp(
        prefix=prefix,
        suffix=suffix,
        dir=dir,
        text=text
    )
    os.close(fd)

    return path


def execute(*args, input=None, ignore_error_code=False):
    class Command:
        def __init__(self, stdout, stderr, returncode):
            self.stdout = stdout
            self.stderr = stderr
            self.returncode = returncode

    log.debug("Execute: {}".format(subprocess.list2cmdline(args)))
    process = subprocess.Popen(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    stdout, stderr = process.communicate(input=input)
    if not ignore_error_code:
        if process.returncode != 0:
            raise RuntimeError("Error executing command",
                               subprocess.list2cmdline(args),
                               process.returncode,
                               stdout,
                               stderr,
                               )

    return Command(stdout, stderr, process.returncode)


def enscript(in_path, out_path, language=None, paper=Paper("a5"), font="Courier8", margins=(40, 40, 40, 40,)):
    ps_path = mktemp(suffix=".ps")

    try:
        args = ["enscript"]

        if language is not None:
            args += ["-E", language]

        if paper is not None:
            args += ["-M", paper.upper]

        if font is not None:
            args += ["-f", font]

        if margins is not None:
            raw_margins = ":".join(str(k) for k in margins)
            args += ["--margins", raw_margins]

        args += [in_path]
        args += ["-o", ps_path]
        execute(*args)

        args = ["ps2pdf"]
        args += [ps_path]
        args += [out_path]
        execute(*args)

    finally:
        os.remove(ps_path)


def pdfbook(in_path, out_path, short_edge=False, frame=False, paper=Paper("a4")):
    args = ["pdfbook"]

    if short_edge:
        args += ["--short-edge"]

    if frame:
        args += ["--frame", "true"]

    args += ["--paper", paper.latex]

    args += [in_path]
    args += ["-o", out_path]

    execute(*args)


def pdfjam(in_path, out_path, selection=None, raw_selection=None, landscape=False, margins=None, paper=Paper("a5")):
    raw_selection = raw_selection or gen_raw_selection(selection)

    args = ["pdfjam"]

    args += [in_path]

    if raw_selection is not None:
        args += [raw_selection]

    if landscape:
        args += ["--landscape"]

    if margins is not None:
        args += ["--trim", " ".join("{}mm".format(-k) for k in margins)]

    args += ["--paper", paper.latex]

    args += ["-o", out_path]

    execute(*args)


def pdfseparate(in_path, out_path_template):
    args = ["pdfseparate"]
    args += [in_path]
    args += [out_path_template]

    execute(*args)


def pdfunite(in_paths, out_path):
    args = ["pdfunite"]
    args += in_paths
    args += [out_path]

    execute(*args)


def pdfselect(in_path, out_path, selection):
    info = pdfinfo(in_path)

    page_path_base = mktemp(suffix="-page")
    page_path_template = page_path_base + "-%d.pdf"

    try:
        pdfseparate(
            in_path=in_path,
            out_path_template=page_path_template,
        )

        selected_pages = [page_path_template % k for k in selection]

        pdfunite(
            in_paths=selected_pages,
            out_path=out_path,
        )
    finally:
        os.remove(page_path_base)

        for i in range(info["Pages"]):
            path = page_path_template % (i + 1,)
            if os.path.exists(path):
                os.remove(path)


def pdfinfo(path):
    command = execute("pdfinfo", path)

    result = {}

    for line in command.stdout.splitlines():
        match = re.match(br"^(?P<key>[^:]*):\s*(?P<value>.*)$", line)
        if match:
            key = match.group("key").decode()
            value = match.group("value").decode()

            if key in [
                "Pages",
                "Page rot",
            ]:  # \d+
                assert re.match(r"^\d+$", value)
                value = int(value)

            if key in [
                "File size",
            ]:  # \d+ bytes
                match = re.match(r"^(?P<value>\d+)\sbytes$", value)
                assert match
                value = int(match.group("value"))

            if key in [
                "Tagged",
                "UserProperties",
                "Suspects",
                "JavaScript",
                "Encrypted",
                "Optimized",
            ]:  # (yes|no)
                assert re.match(r"(yes|no)", value)
                value = True if value == "yes" else False

            if key in [
                "CreationDate",
                "ModDate",
            ]:  # Sat Oct 14 16:50:41 2017 CEST
                value = datetime.datetime.strptime(value, "%c %Z")

            result[key] = value
        else:
            log.warning("PDFInfo line not recognized: {}".format(line))

    assert "Pages" in result

    return result


def lpr(path, printer, double_paged=False):
    args = ["lpr"]
    args += ["-P", printer]
    args += ["-o", "sides={}".format("two-sided-long-edge" if double_paged else "one-sided")]
    args += ["-o", "media=a4"]
    args += [path]

    execute(*args)


class BaseBook:
    def __init__(self, content):
        self.content = content

    def __str__(self):
        return "<{}: {}>".format(type(self).__name__, self.content)

    def __repr__(self):
        return str(self)

    def generate(self, path):
        raise NotImplementedError()


class Book(BaseBook):
    def __init__(self, content):
        super(Book, self).__init__(content=content)

    def generate(self, path):
        content_path = mktemp(suffix=".pdf")

        try:
            res = self.content.generate(content_path)
            assert res == (content_path,)

            pdfbook(
                in_path=content_path,
                out_path=path,
                short_edge=True,
                frame=True,
                paper=Paper("a4"),
            )

        finally:
            os.remove(content_path)

        return path,


class SinglePageBook(BaseBook):
    def __init__(self, content):
        super(SinglePageBook, self).__init__(content=content)

        self.book = Book(self.content)

    def generate(self, path, odd_path=None, even_path=None):
        odd_path = re.sub(r"(.*)(\.[^.]*)", r"\1-odd\2", path)
        even_path = re.sub(r"(.*)(\.[^.]*)", r"\1-even\2", path)

        log.info("Using following paths for SinglePageBook generation: {} {}".format(odd_path, even_path))

        book_path = mktemp(suffix=".pdf")

        try:
            res = self.book.generate(book_path)
            assert res == (book_path,)

            info = pdfinfo(book_path)
            pages_count = info["Pages"]
            assert pages_count % 2 == 0

            pdfselect(book_path, odd_path, selection=range(1, pages_count, 2))
            pdfselect(book_path, even_path, selection=range(pages_count, 1, -2))

        finally:
            os.remove(book_path)

        return odd_path, even_path,


class File:
    def __init__(self, path, selection=None, raw_selection=None, margins=None):
        self.path = path
        self.selection = selection
        self.raw_selection = raw_selection or gen_raw_selection(self.selection)
        self.margins = margins

        if self.margins is not None:
            if not isinstance(self.margins, collections.Iterable):
                self.margins = (self.margins, self.margins, self.margins, self.margins,)

            log.info(type(self.margins))
            log.info(self.margins)

            if len(self.margins) != 4:
                raise ValueError("Wrong length of margins vector: {}".format(len(margins)))

    def __str__(self):
        return "<{}: {}>".format(type(self).__name__, self.path)

    def __repr__(self):
        return str(self)

    def generate(self, path):
        raise NotImplementedError()


class PDFFile(File):
    def __init__(self, path, *args, **kwargs):
        super(PDFFile, self).__init__(path, *args, **kwargs)

    def generate(self, path):
        pdfjam(self.path, path, raw_selection=self.raw_selection, margins=self.margins, paper=Paper("a5"))
        return path,


class CPPFile(File):
    def __init__(self, path, *args, **kwargs):
        super(CPPFile, self).__init__(path, *args, **kwargs)

        if self.raw_selection is not None:
            raise ValueError("Selection is not supported")

    def generate(self, path):
        log.debug("Generate CPP PDF: {}".format(path))

        enscript(
            in_path=self.path,
            out_path=path,
            language="cpp",
            margins=self.margins,
            paper=Paper("a5"),
        )
        return path,


def get_file(path, *args, **kwargs):
    result = None

    if re.match(r".*\.pdf$", path, re.IGNORECASE):
        result = PDFFile(path, *args, **kwargs)
    if re.match(r".*\.(cpp|cxx)$", path, re.IGNORECASE):
        result = CPPFile(path, *args, **kwargs)

    if result is None:
        raise ValueError("Unknown extension", path)
    log.debug("Using: {}".format(result))
    return result


def main(*args):
    parser = argparse.ArgumentParser()
    parser.add_argument("files", nargs="+")
    parser.add_argument("-p", "--print", metavar="PRINTER")
    parser.add_argument("-d", "--double-paged", action="store_true")
    parser.add_argument("-v", "--verbose", action="store_true")
    parser.add_argument("-s", "--select")
    parser.add_argument("-m", "--margins", type=int)

    options = parser.parse_args(args)

    if options.verbose:
        logging.basicConfig(level=logging.DEBUG)

    if options.select is not None and len(options.files) > 1:
        raise ValueError("--select flag is possible for single file")

    book_class = Book if options.double_paged else SinglePageBook

    for path in options.files:
        file = get_file(path, raw_selection=options.select, margins=options.margins)
        book = book_class(file)

        if options.print is None:
            path = re.sub(r"(?:(.*)\.[^.]*|([^.]*))$", r"\1\2-book.pdf", path)
            res = book.generate(path)

            for r in res:
                print(r)
        else:  # print
            res = ()

            try:
                path = mktemp(suffix=".pdf")
                res = book.generate(path)

                for r in res:
                    response = input("print {}? [Y/n] ".format(r))

                    if response in ("n", "N",):
                        continue

                    lpr(
                        path=r,
                        printer=options.print,
                    )

            finally:
                if path not in res:
                    os.remove(path)

                for r in res:
                    os.remove(r)


if __name__ == '__main__':
    main(*sys.argv[1:])
