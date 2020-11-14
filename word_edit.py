# !/usr/bin/python3
# -*- coding: utf-8 -*-
# @Project : Japanese Language Proficiency Test Database Generator
# @Time    : 2020/2/7 19:10
# @Author  : Fang Baole
# @Email   : fbl718@sjtu.edu.cn

from win32com.client import Dispatch
import os


def doc2docx(docPath):
    """Converts a ".doc" file into a ".docx" file

    The function convert a ".doc" file into a ".docx" file in the same forder.

    Args:
        docPath: A string that describes the path of the ".doc" file.

    Returns:
        A string that describes the path of the converted ".docx" file.
    """
    docxPath = docPath + 'x'
    word = Dispatch('Word.Application')
    pathPrefix = os.getcwd() + '\\'
    doc = word.Documents.Open(pathPrefix + docPath)
    doc.SaveAs(pathPrefix + docxPath, FileFormat=12)
    doc.Close()
    word.Quit()
    return docxPath


def auto_num_check(docPath):
    """Converts auto numbering in a word file into plain text.

    The function converts auto numbering in a word file into plain text.

    Args:
        docPath: A string that describes the path of the ".doc" file.

    Returns:
        None
    """
    word = Dispatch('Word.Application')
    pathPrefix = os.getcwd() + '\\'
    doc = word.Documents.Open(pathPrefix + docPath)
    word.Application.Run('ConvertAutoNumToTxt')
    doc.Close()
    word.Quit()
