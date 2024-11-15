'''
----------------------------------------------------------------------------------------------------------------------------------------
Author: Sherif Mehralivand
Email: sherif.mehralivand@ukdd.de
Github: https://github.com/smehralivand
Twitter: @smehralivand
Date: 11/04/2024
----------------------------------------------------------------------------------------------------------------------------------------
'''

from word_converter import doc_to_docx, docx_to_doc, docx_to_txt, txt_to_docx

print('\nMicrosoft Word Converter\n'.upper())
source_path = input('Please enter source directory path: ')
target_path = input('Please enter target path for converted files: ')

while True:

    print('\n1. Convert DOC files into DOCX format.')
    print('\n2. Convert DOCX files into DOC format.')
    print('\n3. Convert DOCX files into TXT format.')
    print('\n4. Convert TXT files into DOCX format.')
    print('\n5. End program.')

    inp = input('\nHow do you want to continue? [1-5] ')

    if inp == '1':
        num = doc_to_docx(source_path,target_path)
        print ('\n{} files were converted\n'.format(num))

    elif inp == '2':
        num = docx_to_doc(source_path, target_path)
        print ('\n{} files were converted\n'.format(num))
    
    elif inp == '3':
        num = docx_to_txt(source_path, target_path)
        print ('\n{} files were converted\n'.format(num))
    
    elif inp == '4':
        num = txt_to_docx(source_path, target_path)
        print ('\n{} files were converted\n'.format(num))

    elif inp == '5':
        break
    
    else:
        print('\nFalse entry. Please try again.\n')
        continue
