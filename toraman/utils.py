import datetime
import re

from html import escape
from lxml import etree

def get_current_time_in_utc():
    return datetime.datetime.utcnow().strftime(r'%Y%M%dT%H%M%SZ')

def html_to_segment(source_or_target_segment, segment_designation, nsmap={'toraman': 'https://cat.toraman.pro'}):

    segment = re.findall(r'<tag[\s\S]+?class="([\s\S]+?)">([\s\S]+?)</tag>|([^<^>]+)',
                        source_or_target_segment)

    segment_xml = etree.Element('{{{0}}}{1}'.format(nsmap['toraman'], segment_designation),
                                nsmap=nsmap)
    for element in segment:
        if element[0]:
            tag = element[0].split()
            tag.insert(0, element[1][len(tag[0]):])
            segment_xml.append(etree.Element('{{{0}}}{1}'.format(nsmap['toraman'], tag[1])))
            segment_xml[-1].attrib['no'] = tag[0]
            if len(tag) > 2:
                segment_xml[-1].attrib['type'] = tag[2]
        elif element[2]:
            segment_xml.append(etree.Element('{{{0}}}text'.format(nsmap['toraman'])))
            segment_xml[-1].text = element[2]

    return segment_xml

def segment_to_html(source_or_target_segment):
    segment_html = ''
    for sub_elem in source_or_target_segment:
        if sub_elem.tag.endswith('}text'):
            segment_html += escape(sub_elem.text)
        else:
            tag = etree.Element('tag')
            tag.attrib['contenteditable'] = 'false'
            tag.attrib['class'] = sub_elem.tag.split('}')[-1]
            tag.text = tag.attrib['class']
            if 'type' in sub_elem.attrib:
                if sub_elem.attrib['type'] == 'beginning' or sub_elem.attrib['type'] == 'end':
                    tag.attrib['class'] += ' ' + sub_elem.attrib['type']
                else:
                    tag.attrib['class'] += ' ' + 'standalone'
            else:
                tag.attrib['class'] += ' ' + 'standalone'

            if 'no' in sub_elem.attrib:
                tag.text += sub_elem.attrib['no']

            segment_html += etree.tostring(tag).decode()

    return segment_html

def segment_to_tm_segment(segment):
    target_segment = ''

    for segment_child in segment:
        if segment_child.tag == '{{{0}}}text'.format(nsmap['toraman']):
            target_segment += segment_child.text
        elif segment_child.tag == '{{{0}}}tag'.format(nsmap['toraman']):
            if segment_child.attrib['type'] == 'beginning':
                target_segment += '<tag{0}>'.format(segment_child.attrib['no'])
            else:
                target_segment += '</tag{0}>'.format(segment_child.attrib['no'])
        else:
            _tag_label = segment_child.tag.split('}')[1]
            if 'no' in segment_child.attrib:
                _tag_label += segment_child.attrib['no']
            target_segment += '<{0}/>'.format(_tag_label)

    return target_segment
