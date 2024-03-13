from pptx import Presentation
from .utilities import fill_BOM, file_type
# import argparse

# parser = argparse.ArgumentParser()
# parser.add_argument('-i', '--input', type= file_type)
# parser.add_argument('-o', '--output', type= file_type)

# prs = Presentation(parser.input)

prs = Presentation('data/IT-prueva.pptx')
task_structure = [('010', 'tareas previas', 1,4), ('020', 'desmontaje',5,8)]

fill_BOM(prs)

# for task in task_structure:
#     write_task_number_text (prs, *task)

# task_dict = get_task_dict(prs)

# Numbering the paragraphs:

# for key in task_dict:
#     step_paragraphs = get_step_paragraphs(task_dict[key])
#     number_paragraphs(step_paragraphs)

prs.save(parser.output)
