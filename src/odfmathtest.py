#!/usr/bin/env python
# -*- coding: utf-8 -*-
from xml.dom.minidom import parseString
from xml.dom import Node
import odf
import odf.opendocument
import odf.text
from odf.element import Element
from namespaces import MATHNS
math_templ = u'\
<math xmlns="http://www.w3.org/1998/Math/MathML">\
<semantics>\
<annotation encoding="StarMath 5.0">%s</annotation>\
</semantics></math>'
def gen_odf_math_(parent):
    elem = Element(qname = (MATHNS,parent.tagName))
    if parent.attributes:
        for attr, value in parent.attributes.items():
            elem.setAttribute((MATHNS,attr), value, check_grammar=False)
    for child in parent.childNodes:
        if child.nodeType == Node.TEXT_NODE:
            text = child.nodeValue
            elem.addText(text, check_grammar=False)
        else:
            elem.addElement(gen_odf_math_(child), check_grammar=False)
    return elem
def gen_odf_math(starmath_string):
    u'''
    Generating odf.math.Math element
    '''
    mathml = math_templ % (starmath_string)
    math_ = parseString(mathml.encode('utf-8'))
    math_ = math_.documentElement
    odf_math = gen_odf_math_(math_)
    return odf_math
def main():
    doc = odf.opendocument.OpenDocumentText()
    p = odf.text.P(text=u'text')
    df = odf.draw.Frame( zindex=0, anchortype='as-char')
    p.addElement(df)
    doc.text.addElement(p)
    formula = 'c = sqrt(a^2+b_2) + %ialpha over %ibeta'
    math = gen_odf_math(formula)
    do = odf.draw.Object()
    do.addElement(math)
    df.addElement(do)
    outputfile = u'result'
    doc.save(outputfile, True)
if __name__ == '__main__':
    main()
