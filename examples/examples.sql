DECLARE
    CURSOR c_bill IS
        SELECT 'Bread' as name, 'pcs' as unit, 2 as q, 1.2 as price FROM dual UNION ALL
        SELECT 'Milk' as name, 'pcs' as unit, 1 as q, 0.8 as price FROM dual UNION ALL
        SELECT 'Eggs' as name, 'pack' as unit, 3 as q, 2.5 as price FROM dual UNION ALL
        SELECT 'Butter' as name, 'pcs' as unit, 1 as q, 3 as price FROM dual UNION ALL
        SELECT 'Cookies' as name, 'pcs' as unit, 1 as q, 2 as price FROM dual;

    lnDok pls_integer;
    lnParagraph pls_integer;
    lnNum pls_integer;
    lnBreak pls_integer;
    lnHeader pls_integer;
    lnTable pls_integer;
    lnCounter pls_integer := 1;
    lnSum number := 0;
    lrPage ZT_WORD.r_page;
    
    lbDokument blob;

BEGIN
    --first we create a new document and get a reference ID
    --function argument is document author name 
    lnDok := ZT_WORD.f_new_document('Alan Ford');
    
    
    --then we create document elements (paragraphs with text, tables, lists...)
    --we always reference with document ID
    lnParagraph := ZT_WORD.f_new_paragraph(
        p_doc_id => lnDok,
        p_alignment_h => 'CENTER',
        p_text => 'HELLO WORLD',
        p_font => 
            ZT_WORD.f_font(
                p_font_name => 'Times new Roman',
                p_font_size => 28,
                p_bold => true
            )
        );


    --empty paragraph without any text
    lnParagraph := ZT_WORD.f_new_paragraph(
        p_doc_id => lnDok
            );

            
    --here we add various text to paragraph, which we first create with some initial text
    --we reference to paragraph with document ID and paragraph ID together
    lnParagraph := ZT_WORD.f_new_paragraph(
        p_doc_id => lnDok,
        p_text => 'This is an example '
            );

    ZT_WORD.p_add_text(
        p_doc_id => lnDok,
        p_paragraph_id => lnParagraph,
        p_text => 'of bold text ',
        p_font => ZT_WORD.f_font(p_bold => true)
        );

    ZT_WORD.p_add_text(
        p_doc_id => lnDok,
        p_paragraph_id => lnParagraph,
        p_text => 'and then italic text ',
        p_font => ZT_WORD.f_font(p_italic => true)
        );

    ZT_WORD.p_add_text(
        p_doc_id => lnDok,
        p_paragraph_id => lnParagraph,
        p_text => 'also underline ',
        p_font => ZT_WORD.f_font(p_underline => true)
        );

    ZT_WORD.p_add_text(
        p_doc_id => lnDok,
        p_paragraph_id => lnParagraph,
        p_text => 'and red.',
        p_font => ZT_WORD.f_font(p_color => 'FF0000')
        );
    

    --now we add next paragraph with heading style
    --we also have bullets and numbering example
    lnParagraph := ZT_WORD.f_new_paragraph(
        p_doc_id => lnDok,
        p_text => 'Bullets and numbering',
        p_style => 'Heading1'
            );

    lnParagraph := ZT_WORD.f_new_paragraph(
        p_doc_id => lnDok,
        p_text => 'Bullets example',
        p_style => 'Heading2'
            );

    lnNum := ZT_WORD.f_new_bullet(
        p_doc_id => lnDok,
        p_char => '·',
        p_font => 'Symbol');

    FOR t IN 1 .. 10 LOOP
        lnParagraph := ZT_WORD.f_new_paragraph(
            p_doc_id => lnDok,
            p_list_id => lnNum,
            p_text => 'Bulleting line ' || t
                );
    END LOOP;


    lnParagraph := ZT_WORD.f_new_paragraph(
        p_doc_id => lnDok,
        p_text => 'Numbering example',
        p_style => 'Heading2'
            );
    
    lnNum := ZT_WORD.f_new_numbering(p_doc_id => lnDok);
    
    FOR t IN 1 .. 10 LOOP
        lnParagraph := ZT_WORD.f_new_paragraph(
            p_doc_id => lnDok,
            p_list_id => lnNum,
            p_text => 'Numbering line ' || t
                );
    END LOOP;


    lnParagraph := ZT_WORD.f_new_paragraph(
        p_doc_id => lnDok,
        p_text => 'And further we go... ',
        p_style => 'Heading1'
            );

    lnParagraph := ZT_WORD.f_new_paragraph(
        p_doc_id => lnDok,
        p_text => '...to another page'
            );


    --a section break with new header
    --first we create a header and add content in it
    --then we define a page object with header and make a section break
    lnHeader := ZT_WORD.f_new_container(
        p_doc_id => lnDok, 
        p_type => 'HEADER');

    lnParagraph := ZT_WORD.f_new_paragraph(
        p_doc_id => lnDok,
        p_container_id => lnHeader,
        p_alignment_h => 'CENTER',
        p_text => 'New section header for table example',
        p_font => 
            ZT_WORD.f_font(
                p_font_name => 'Times new Roman',
                p_font_size => 18,
                p_bold => true,
                p_underline => true
            )
        );

    lrPage := ZT_WORD.f_get_default_page(p_doc_id => lnDok);
    lrPage.header_ref := lnHeader;

    lnBreak := ZT_WORD.f_new_section_break(
        p_doc_id => lnDok,
        p_section_type => 'next_page',
        p_page => lrPage);

    lnParagraph := ZT_WORD.f_new_paragraph(
        p_doc_id => lnDok,
        p_text => 'Here we are with bill example...'
            );


    --tables and table cell manipulation
    lnTable := ZT_WORD.f_new_table(
        p_doc_id => lnDok, 
        p_rows => 7,
        p_columns => 5,
        p_columns_width => '5000, 1000, 1000, 1000, 1000');
    
    ZT_WORD.p_table_cell(
        p_doc_id => lnDok, 
        p_table_id => lnTable, 
        p_row => 1, 
        p_column => 1, 
        p_alignment_h => 'LEFT',
        p_background_color => 'CCCCCC',
        p_text => 'Item');
    ZT_WORD.p_table_cell(p_doc_id => lnDok, p_table_id => lnTable, p_row => 1, p_column => 2, p_alignment_h => 'CENTER', p_background_color => 'CCCCCC', p_text => 'Unit');
    ZT_WORD.p_table_cell(p_doc_id => lnDok, p_table_id => lnTable, p_row => 1, p_column => 3, p_alignment_h => 'RIGHT', p_background_color => 'CCCCCC', p_text => 'Quantity');
    ZT_WORD.p_table_cell(p_doc_id => lnDok, p_table_id => lnTable, p_row => 1, p_column => 4, p_alignment_h => 'RIGHT', p_background_color => 'CCCCCC', p_text => 'Price');
    ZT_WORD.p_table_cell(p_doc_id => lnDok, p_table_id => lnTable, p_row => 1, p_column => 5, p_alignment_h => 'RIGHT', p_background_color => 'CCCCCC', p_text => 'Total');

    FOR t in c_bill LOOP
        lnCounter := lnCounter + 1;
        ZT_WORD.p_table_cell(p_doc_id => lnDok, p_table_id => lnTable, p_row => lnCounter, p_column => 1, p_alignment_h => 'LEFT', p_text => t.name);
        ZT_WORD.p_table_cell(p_doc_id => lnDok, p_table_id => lnTable, p_row => lnCounter, p_column => 2, p_alignment_h => 'CENTER', p_text => t.unit);
        ZT_WORD.p_table_cell(p_doc_id => lnDok, p_table_id => lnTable, p_row => lnCounter, p_column => 3, p_alignment_h => 'RIGHT', p_text => to_char(t.q, 'FM999G990D00') );
        ZT_WORD.p_table_cell(p_doc_id => lnDok, p_table_id => lnTable, p_row => lnCounter, p_column => 4, p_alignment_h => 'RIGHT', p_text => to_char(t.price, 'FM999G990D00') || '€');
        ZT_WORD.p_table_cell(p_doc_id => lnDok, p_table_id => lnTable, p_row => lnCounter, p_column => 5, p_alignment_h => 'RIGHT', p_text => to_char(t.q * t.price, 'FM999G990D00') || '€');
        
        lnSum := lnSum + (t.q * t.price);
    END LOOP; 

    ZT_WORD.P_TABLE_MERGE_CELLS(lnDok, lnTable, 7, 1, 7, 4);

    ZT_WORD.p_table_cell(
        p_doc_id => lnDok, 
        p_table_id => lnTable, 
        p_row => 7, 
        p_column => 1, 
        p_alignment_h => 'LEFT', 
        p_background_color => 'CCCCCC',
        p_text => 'Total');

    ZT_WORD.p_table_cell(
        p_doc_id => lnDok, 
        p_table_id => lnTable, 
        p_row => 7, 
        p_column => 5, 
        p_alignment_h => 'RIGHT', 
        p_background_color => 'CCCCCC',
        p_text => to_char(lnSum, 'FM999G990D00') || '€');
    

    --at the end we finish the document
    --function returns the document as blob variable
    lbDokument := ZT_WORD.f_make_document(lnDok);
    
    --for testing purposes we save the document in directory
    ZT_WORD.p_save_file(
        p_document => lbDokument,
        p_folder => 'D_SHARED',
        p_file_name => 'ZT_WORD example.docx');
    
END;