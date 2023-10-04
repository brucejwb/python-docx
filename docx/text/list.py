# encoding: utf-8

"""
The |ListParagraph| object and related proxy classes.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import random
from docx.text.paragraph import Paragraph
from docx.shared import Inches


NUMBERING_FORMAT_TO_ABSTRACT_NUM_ID = {
    "closedDiamond": 0,
    "upperRoman": 1,
    "openDiamond": 2,
    "decimal": 3,
    "lowerLetter": 4,
    "lowerRoman": 5,
    "bullet": 6,
    "arrow": 7,
    "star": 8,
    "upperLetter": 9,
}

class ListParagraph(object):
    """
    Proxy object for controlling a set of ``<w:p>`` grouped together in a list.
    """
    
    def create_num(self, level, numbering_format):
        self._numbering = self.document.part.numbering_part.numbering_definitions._numbering
        num = self._numbering.add_num(NUMBERING_FORMAT_TO_ABSTRACT_NUM_ID.get(numbering_format, 2))
        for i in range(level+1):
            override = num.add_lvlOverride(i)
            override.add_startOverride(1)
        return num
    
    def __init__(self, parent, document, num_id, numbering_format="decimal", level=0):
        self._parent = parent
        self.document = document
        self.level = level
        
        if not num_id:
            num = self.create_num(level, numbering_format)
            self.numId = num.numId
        else:
            self.numId = num_id

    def add_list(self, num_id=None, numbering_format="decimal"):
        """
        Add a list indented one level below the current one, having a paragraph
        style *style*. Note that the document will only be altered once the
        first item has been added to the list.
        """
        return ListParagraph(
            self._parent,
            self.document,
            num_id=num_id,
            numbering_format=numbering_format,
            level=self.level+1,
        )
    
    def add_paragraph(self, text=None):
        p = self.document._body.add_paragraph(text)
        p.paragraph_format.left_indent = Inches(.25*(self.level))
        return p
    
    def insert_paragraph_before(self, paragraph, text):
        p = paragraph.insert_paragraph_before(text=text)
        p.paragraph_format.left_indent = Inches(.25*(self.level))
        return p
    
    def add_item(self, text=None):
        """
        Add a paragraph item to the current list, having text set to *text* and
        a paragraph style *style*
        """
        p = self.document._body.add_paragraph(text)
        p.level = self.level
        p.numId = self.numId
        return p
    
    def insert_item_before(self, paragraph, text=None):
        """
        Add a paragraph item to the current list, before the inputted paragraph.
        """
        p = paragraph.insert_paragraph_before(text=text)
        p.level = self.level
        p.numId = self.numId
        return p

    @property
    def items(self):
        """
        Sequence of |Paragraph| instances corresponding to the item elements
        in this list paragraph.
        """
        return [paragraph for paragraph in self._parent.paragraphs
                if paragraph.numId == self.numId]
