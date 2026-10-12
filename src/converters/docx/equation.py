"""Convert Feishu LaTeX equations into editable Word math objects."""

from lxml import etree


_MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'


def latex_to_omml(latex):
    r"""Return an ``m:oMath`` element for a Feishu equation's LaTeX source.

    mathml2omml 0.0.2 emits an incorrect closing tag for stretchy accents
    such as ``\bar{x}``. Repair that known output before parsing it. Keeping
    the conversion here also lets the caller retain the source text if a
    different LaTeX construct cannot be converted.
    """
    from latex2mathml.converter import convert as latex_to_mathml
    from mathml2omml import convert as mathml_to_omml

    mathml = latex_to_mathml(latex)
    omml = mathml_to_omml(mathml)
    for position in ('top', 'bot'):
        omml = omml.replace(
            f'<m:pos m:val="{position}"/></m:groupChr>',
            f'<m:pos m:val="{position}"/></m:groupChrPr>',
        )

    wrapper = etree.fromstring(
        f'<root xmlns:m="{_MATH_NS}">{omml}</root>'.encode('utf-8')
    )
    if len(wrapper) != 1 or wrapper[0].tag != f'{{{_MATH_NS}}}oMath':
        raise ValueError('MathML converter did not return an Office math element')
    equation = wrapper[0]
    # The converter writes empty n-ary limits as visible placeholders. Word
    # math requires the limit elements, so hide an unused limit explicitly.
    for nary in equation.iter(f'{{{_MATH_NS}}}nary'):
        properties = nary.find(f'{{{_MATH_NS}}}naryPr')
        if properties is None:
            continue
        for limit_name, hide_name in (('sub', 'subHide'), ('sup', 'supHide')):
            limit = nary.find(f'{{{_MATH_NS}}}{limit_name}')
            if limit is not None and len(limit) == 0 and not (limit.text or '').strip():
                hide = etree.SubElement(properties, f'{{{_MATH_NS}}}{hide_name}')
                hide.set(f'{{{_MATH_NS}}}val', '1')
    return equation
