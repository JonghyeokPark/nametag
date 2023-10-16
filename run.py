from pptx import Presentation
from pptx.util import Inches

prs = Presentation('template.pptx')

import copy           
def copy_slide(prs, index):
    template = prs.slides[index]
    try:
        blank_slide_layout = prs.slide_layouts.get_by_name()
    except:
        blank_slide_layout = prs.slide_layouts[0]

    copied_slide = prs.slides.add_slide(blank_slide_layout)
    
    for shape in template.shapes:
        elem = shape.element
        new_elem = copy.deepcopy(elem)
        copied_slide.shapes._spTree.insert_element_before(new_elem, 'p:extLst')

    return copied_slide

def main():
    print("Welcom Nametag")
    copied_slide = copy_slide(prs, 0)
    prs.save('output.pptx')

if __name__ == "__main__":
    main()