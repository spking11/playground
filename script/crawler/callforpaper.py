import requests
from bs4 import BeautifulSoup, Tag
import xlsxwriter
from xlsxwriter import worksheet
from urllib.parse import urljoin
from collections import defaultdict
from functools import reduce
import os
import requests_cache

requests_cache.install_cache(cache_name="cache", backend="sqlite")


class Conference:
    def __init__(self, name: str, name_link: str, date: str) -> None:
        self.name = name
        self.name_link = name_link
        self.date = date


class Journal(Conference):
    pass


class CCFList:
    def __init__(self, linkbase: str) -> None:
        self.conferences: dict[str, list[Conference]] = defaultdict(list)
        self.journals: dict[str, list[Journal]] = defaultdict(list)
        self.linkbase = linkbase

    def resolveLink(self, link: str):
        return urljoin(self.linkbase, link)

    def setField(self, tag: Tag):
        self.field = tag.text
        self.field_link = self.resolveLink(tag.attrs["href"])

    def addConference(self, rank: str, tag: Tag):
        name_tag = tag.select_one(".table-tr-name a")
        self.conferences[rank].append(
            Conference(
                name_tag.text.strip(),
                self.resolveLink(name_tag.attrs["href"]),
                tag.select_one(".table-tr-date").contents[0].text.strip(),
            )
        )

    def addJournal(self, rank: str, tag: Tag):
        # print(tag.select_one(".table-tr-jname").contents[0].text.strip())
        name_tag = tag.select_one(".table-tr-si a")
        self.journals[rank].append(
            Journal(
                name_tag.text.strip(),
                self.resolveLink(name_tag.attrs["href"]),
                tag.select_one(".table-tr-date").contents[0].text.strip(),
            )
        )

    def getConferenceCount(self):
        return reduce(lambda x, value: x + len(value), self.conferences.values(), 0)

    def getJournalCount(self):
        return reduce(lambda x, value: x + len(value), self.journals.values(), 0)


ccflists: list[CCFList] = []


def fetchData():
    for i in range(1, 11):
        url = f"http://123.57.137.208/ccf/ccf-{i}.jsp"
        response = requests.get(url)
        if response.status_code == 200:
            ccflist = CCFList(url)
            ccflists.append(ccflist)

            soup = BeautifulSoup(response.text, "html.parser")
            ccflist.setField(soup.select_one(".field-mark > a"))
            # [conferences, journals] = soup.select(".ccf-frame h2:not([style])")
            ranks = soup.select("h3")
            for i, rank in enumerate(ranks):
                table = rank.find_next(None, {"class": "ccf-table"})
                for row in table.select(".table-tr-content:not(:first-child)"):
                    if i < 3:
                        ccflist.addConference(rank.text, row)
                    else:
                        ccflist.addJournal(rank.text, row)
        else:
            print(f"请求失败，状态码: {response.status_code}")


def writeMerge(
    worksheet: worksheet.Worksheet,
    row: int,
    col: int,
    size: int,
    data: str,
    link: str | None = None,
):
    if size == 0:
        return
    if size == 1:
        if link:
            worksheet.write_url(row, col, link, string=data)
        else:
            worksheet.write(row, col, data)
    else:
        worksheet.merge_range(
            row,
            col,
            row + size - 1,
            col,
            data,
        )
        if link:
            worksheet.write_url(row, col, link, string=data)


def writeData():
    with xlsxwriter.Workbook(os.path.splitext(__file__)[0] + ".xlsx") as workbook:
        worksheet = workbook.add_worksheet()
        worksheet.write_row("A1", ("专业领域", "种类", "rank", "名称", "截稿日期"))
        row_num = 1
        for ccflist in ccflists:
            writeMerge(
                worksheet,
                row_num,
                0,
                ccflist.getConferenceCount() + ccflist.getJournalCount(),
                ccflist.field,
                ccflist.field_link,
            )
            writeMerge(
                worksheet,
                row_num,
                1,
                ccflist.getConferenceCount(),
                "中国计算机学会推荐国际学术会议",
            )
            writeMerge(
                worksheet,
                row_num + ccflist.getConferenceCount(),
                1,
                ccflist.getJournalCount(),
                "中国计算机学会推荐国际学术刊物",
            )
            for rank, conferences in ccflist.conferences.items():
                writeMerge(worksheet, row_num, 2, len(conferences), rank)
                for conference in conferences:
                    worksheet.write_url(
                        row_num, 3, conference.name_link, string=conference.name
                    )
                    worksheet.write(row_num, 4, conference.date)
                    row_num += 1

            for rank, journals in ccflist.journals.items():
                writeMerge(worksheet, row_num, 2, len(journals), rank)
                for journal in journals:
                    worksheet.write_url(
                        row_num, 3, journal.name_link, string=journal.name
                    )
                    worksheet.write(row_num, 4, journal.date)
                    row_num += 1


fetchData()
writeData()
