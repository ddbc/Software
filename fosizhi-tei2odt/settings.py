#-*- coding:utf-8 -*-

# default - can be over-ridden by a command switch
TEI_BASE = 'fosizhi/xml/'

# tags to pass over (but process text/child-nodes)
PASS_TAGS = [
    'div',
    'list',
    'lg',
]

# tags to ignore (don't render at all)
IGNORE_TAGS = [
    'pb',
    'choice',
    'sic',
    'corr',
    'reg',
    'orig',
    'gap',
    'lb',
    'unclear',
    'space',
]

# tags to render in character styles
CHARSTYLE_TAGS = [
    'date',
    'persName',
    'placeName',
    'roleName',
    'c',
    'seg',
]

# tags to render in paragraph styles
PARASTYLE_TAGS = [
    'p',
    'head',
    'byline',
    'item',
    'l',
    'closer',
]

# tags to render as footnotes
FOOTNOTE_TAGS = [
    'note',
]


# TEI-tag processing functions begin
####################################

def gap(writer, el, stack):
    extent = el.get('extent')
    reason = el.get('reason')
    if extent is None or reason is None:
        return False
    charStyle = 'gap_{0}'.format(reason)
    # insert U+303F, "〿"
    writer.write_string(u'\u303f' * int(extent), charStyle)
    return True


def space(writer, el, stack):
    quantity = el.get('quantity')
    if quantity is None:
        return False
    # insert U+3000, full-width space
    writer.write_string(u'\u3000' * int(quantity))
    return True


def unclear(writer, el, stack):
    if len(el) > 0 or len(el.text) != 1:
        return False
    writer.write_string(u'{0}（？）'.format(el.text), 'unclear')
    return True


def pb(writer, el, stack):
    pageNo = el.get('n')
    if pageNo is None:
        return False
    writer.write_string('[p{0}]'.format(pageNo), 'pageNo')
    return True


def choice(writer, el, stack):

    if el.find('sic') is not None and el.find('corr') is not None:
        charStyle = 'corr'
        mainText = el.find('corr').text
        fnText = u'Corrected by DDBC - source text has 「{0}」'\
                    .format(el.find('sic').text)

    elif el.find('orig') is not None and el.find('reg') is not None:
        charStyle = 'reg'
        mainText = el.find('reg').text
        fnText = u'Regularized by DDBC - source text has 「{0}」'\
                        .format(el.find('orig').text)

    else:
        return False

    writer.write_string(mainText, charStyle)
    writer.insert_footnote(fnText)
    return True


def date(writer, el, stack):

    if el.get('notBefore') is not None and el.get('notAfter') is not None:
        westernDate = '{0} -> {1}'.format(
                            el.get('notBefore'), el.get('notAfter'))
    elif el.get('from') is not None and el.get('to') is not None:
        westernDate = '{0} -> {1}'.format(el.get('from'), el.get('to'))
    elif el.get('when') is not None:
        westernDate = el.get('when')
    else:
        return False

    writer.write_string('({0})'.format(westernDate), 'westernDate')
    return True


def lb(writer, el, stack):
    writer.close_para()
    return True


def persName(writer, el, stack):
    aid = el.get('key')
    appears_as = joinText(el)
    page = writer.view_cursor.Page
    personIndex.write(
        (u'"{0}","{1}","{2}"\n'.format(aid, appears_as, page)).encode('utf-8'))
    return True


def placeName(writer, el, stack):
    aid = el.get('key')
    appears_as = joinText(el)
    page = writer.view_cursor.Page
    placeIndex.write(
        (u'"{0}","{1}","{2}"\n'.format(aid, appears_as, page)).encode('utf-8'))
    return True



# some prep for the index files
# (executed when the module is imported)
import atexit

personIndex = open('personIndex.csv', 'w')
personIndex.write('"aid","appears_as","page"\n')
placeIndex = open('placeIndex.csv', 'w')
placeIndex.write('"aid","appears_as","page"\n')


def closeup():
    personIndex.close()
    placeIndex.close()

atexit.register(closeup)


# helper function
def joinText(el):
    text = (el.text, '')[not el.text]
    for subEl in el:
        text += joinText(subEl) + (subEl.tail, '')[not subEl.tail]
    return text
