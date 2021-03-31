CREATE OR REPLACE PACKAGE zt_word AS
/******************************************************************************
    Author:     Zoran Tica
                ZT-TECH, racunalniške storitve s.p.
                http://www.zt-tech.eu
    
    PURPOSE:    A package for Microsoft Word DOCX documents generation 

    REVISIONS:
    Ver        Date        Author           Description
    ---------  ----------  ---------------  ------------------------------------
    0.1        28/10/2016  Zoran Tica       1. Created this package.
    1.0        15/10/2017  Zoran Tica       2. First public version.
    2.0        30/03/2020  Zoran Tica       3. Images; table borders


    ----------------------------------------------------------------------------
    Copyright (C) 2017 - Zoran Tica

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in
    all copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
    THE SOFTWARE.
    ----------------------------------------------------------------------------
*/

TYPE t_table_number IS TABLE OF number;
TYPE t_table_vc2 IS TABLE OF varchar2(32000);


--document elements definition
TYPE r_font IS RECORD (
    from_paragraph boolean,
    font_name varchar2(50),
    font_size number,
    bold boolean,
    italic boolean,
    underline boolean,
    color varchar2(6)
    );

TYPE r_text IS RECORD (
    text varchar2(32000),
    font r_font,
    image_data zt_word.r_image_data
    );
TYPE t_text IS TABLE OF r_text;

TYPE r_paragraph IS RECORD (
    alignment_h varchar2(10),
    space_before number,
    space_after number,
    style varchar2(100),
    list_id pls_integer,
    texts t_text 
    );
TYPE t_paragraphs IS TABLE OF r_paragraph;   

TYPE r_border IS RECORD (
    border_type varchar2(50),
    width number,
    color varchar2(6)
    );

TYPE r_table_cell IS RECORD (
    alignment_v varchar2(20),
    border_top r_border,
    border_bottom r_border,
    border_left r_border,
    border_right r_border,
    merge_v varchar2(50),
    merge_h pls_integer,
    background_color varchar2(6),
    paragraph r_paragraph 
    );
TYPE t_table_cells IS TABLE OF r_table_cell INDEX BY varchar2(10);

TYPE r_table IS RECORD (
    width pls_integer,
    rows_num pls_integer,
    columns_num pls_integer,
    column_width t_table_number,
    cells t_table_cells,
    border_top r_border,
    border_bottom r_border,
    border_left r_border,
    border_right r_border,
    border_inside_h r_border,
    border_inside_v r_border
    );

TYPE r_list IS RECORD (
    list_type varchar2(20),
    num_start_value varchar2(10),
    bullet_char varchar2(10),
    bullet_font varchar2(50)
    );
TYPE t_list IS TABLE OF r_list;


TYPE r_page IS RECORD (
    margin_top number,
    margin_bottom number,
    margin_left number,
    margin_right number,
    header_h number,
    footer_h number,
    header_ref pls_integer,
    footer_ref pls_integer,
    orientation varchar2(10),
    width number,
    height number);

TYPE r_break IS RECORD (
    break_type varchar2(20),
    section_type varchar2(20),
    page r_page);


TYPE r_image_data IS RECORD (
    image_id pls_integer,
    width number,
    height number,
    rotate_angle number,
    extent_area_left number,
    extent_area_top number,
    extent_area_right number,
    extent_area_bottom number,
    inline_yn varchar2(1),
    relative_from_h varchar2(10),
    position_type_h varchar2(10),
    position_align_h varchar2(10),
    position_h number,
    relative_from_v varchar2(10),
    position_type_v varchar2(10),
    position_align_v varchar2(10),
    position_v number
    );
TYPE r_image IS RECORD (
    image_name varchar2(50),
    image_file blob,
    rel_id pls_integer
    );
TYPE t_images IS TABLE OF r_image;


TYPE r_element IS RECORD (
    element_type varchar2(20),
    paragraph r_paragraph,
    table_data r_table,
    break_data r_break,
    image_data r_image_data
    );
TYPE t_elements IS TABLE OF r_element;

TYPE r_container IS RECORD (
    container_type varchar2(20),
    rel_id pls_integer,
    elements t_elements
    );
TYPE t_containers IS TABLE OF r_container;

TYPE r_document IS RECORD (
    author varchar2(500),
    create_date date,
    default_page r_page,
    unit varchar2(20),
    lang varchar2(20),
    rels_id pls_integer,
    containers t_containers,
    images t_images,
    lists t_list);
TYPE t_documents IS TABLE OF r_document;




--public functions
FUNCTION f_new_document(
    p_author varchar2 default null,
    p_default_page r_page default null,
    p_unit varchar2 default 'cm',
    p_lang varchar2 default 'en-US') RETURN pls_integer;

FUNCTION f_new_paragraph(
    p_doc_id number,
    p_container_id pls_integer default null,
    p_alignment_h varchar2 default 'LEFT',  --'left', 'right', 'center', 'both' 
    p_space_before number default 0,
    p_space_after number default 0,
    p_style varchar2 default null,
    p_list_id pls_integer default null,
    p_text varchar2 default null,
    p_font r_font default null
    ) RETURN pls_integer;


FUNCTION f_border(
    p_border_type varchar2 default null,  --look at the end of the package for list of possible values
    p_width number default null,
    p_color varchar2 default null --RRGGBB hex format
    ) RETURN r_border;

PROCEDURE p_table_cell(
    p_doc_id number,
    p_table_id pls_integer,
    p_row pls_integer,
    p_column pls_integer,
    p_alignment_h varchar2 default 'LEFT',  --'left', 'right', 'center', 'both'
    p_alignment_v varchar2 default 'TOP',  --'top', 'center', 'bottom'
    p_border_top r_border default null,
    p_border_bottom r_border default null,
    p_border_left r_border default null,
    p_border_right r_border default null,
    p_background_color varchar2 default null, --RRGGBB hex format
    p_text varchar2 default null,
    p_font r_font default null,
    p_image_data r_image_data default null
    );

PROCEDURE p_table_merge_cells(
    p_doc_id number,
    p_table_id pls_integer,
    p_from_row pls_integer,
    p_from_column pls_integer,
    p_to_row pls_integer,
    p_to_column pls_integer);

FUNCTION f_new_table(
    p_doc_id number,
    p_rows pls_integer,
    p_columns pls_integer,
    p_table_width pls_integer default null,
    p_columns_width varchar2 default null,  --comma separated string
    p_border_top r_border default null,
    p_border_bottom r_border default null,
    p_border_left r_border default null,
    p_border_right r_border default null,
    p_border_inside_h r_border default null,
    p_border_inside_v r_border default null
    ) RETURN pls_integer;

FUNCTION f_new_numbering(
    p_doc_id number,
    p_start_value varchar2 default '1') RETURN pls_integer;

FUNCTION f_new_bullet(
    p_doc_id number,
    p_char varchar2 default 'o',
    p_font varchar2 default 'Courier New') RETURN pls_integer;

FUNCTION f_new_page_break(p_doc_id number) RETURN pls_integer;

FUNCTION f_new_section_break(
    p_doc_id number,
    p_section_type varchar2,  --'nextPage', 'oddPage', 'evenPage' 
    p_page_template varchar2 default null,
    p_page r_page default null
    ) RETURN pls_integer;


FUNCTION f_image_data(
    p_image_id pls_integer,
    p_width number default 0,
    p_height number default 0,
    p_rotate_angle number default 0,
    p_extent_area_left number default 0,
    p_extent_area_top number default 0,
    p_extent_area_right number default 0,
    p_extent_area_bottom number default 0,
    p_inline_yn varchar2 default 'Y',
    p_relative_from_h varchar2 default 'page',  --page, margin, column
    p_position_type_h varchar2 default 'align',  --align, posOffset
    p_position_align_h varchar2 default 'center',  --left, right, center
    p_position_h number default 0,  --offset from object
    p_relative_from_v varchar2 default 'page',
    p_position_type_v varchar2 default 'align',
    p_position_align_v varchar2 default 'center',  --top, bottom, center
    p_position_v number default 0
    ) RETURN r_image_data;

/*funtion adds image to document and returns image ID - it doesn't insert image instance on pages*/
FUNCTION f_add_image_to_document(
    p_doc_id number,
    p_filename varchar2,
    p_image blob
    ) RETURN pls_integer;

/*procedure inserts image instance on document pages*/
FUNCTION f_new_image_instance(
    p_doc_id number,
    p_container_id pls_integer default null,
    p_image_data r_image_data
    ) RETURN pls_integer;




PROCEDURE p_set_default_page(
    p_doc_id number,
    p_page r_page);

FUNCTION f_get_default_page(
    p_doc_id number) RETURN r_page;

FUNCTION f_get_page(
    p_doc_id number default null,
    p_width number,
    p_height number,
    p_margin_top number,
    p_margin_bottom number,
    p_margin_left number,
    p_margin_right number,
    p_header_height number,
    p_footer_height number,
    p_header_ref pls_integer default null,
    p_footer_ref pls_integer default null,
    p_orientation varchar2  --'portrait' or 'landscape'
    ) RETURN r_page;
    


FUNCTION f_font(
    p_from_paragraph boolean default false,
    p_font_name varchar2 default 'Calibri',
    p_font_size number default 11,
    p_bold boolean default false,
    p_italic boolean default false,
    p_underline boolean default false,
    p_color varchar2 default '000000'  --RRGGBB hex format
    ) RETURN r_font;

PROCEDURE p_add_text(
    p_doc_id number,
    p_container_id pls_integer default null,
    p_paragraph_id number,
    p_text varchar2 default null,
    p_font r_font default null,
    p_image_data r_image_data default null
    );


FUNCTION f_new_container(
    p_doc_id number,
    p_type varchar2 default 'DOCUMENT'  --'DOCUMENT', 'HEADER', 'FOOTER'
    ) RETURN pls_integer;


FUNCTION f_make_document(
    p_doc_id number) RETURN blob;



--mostly for testing purposes
--procedure saves a document (or some other blob) into a file 
PROCEDURE p_save_file(
    p_document blob,
    p_file_name varchar2 default 'my_document.docx',
    p_folder varchar2 default 'MY_FOLDER'
    );



/*
possible values for cell border line type:
    single - a single line
    dashDotStroked - a line with a series of alternating thin and thick strokes
    dashed - a dashed line
    dashSmallGap - a dashed line with small gaps
    dotDash - a line with alternating dots and dashes
    dotDotDash - a line with a repeating dot - dot - dash sequence
    dotted - a dotted line
    double - a double line
    doubleWave - a double wavy line
    inset - an inset set of lines
    nil - no border
    none - no border
    outset - an outset set of lines
    thick - a single line
    thickThinLargeGap - a thick line contained within a thin line with a large-sized intermediate gap
    thickThinMediumGap - a thick line contained within a thin line with a medium-sized intermediate gap
    thickThinSmallGap - a thick line contained within a thin line with a small intermediate gap
    thinThickLargeGap - a thin line contained within a thick line with a large-sized intermediate gap
    thinThickMediumGap - a thick line contained within a thin line with a medium-sized intermediate gap
    thinThickSmallGap - a thick line contained within a thin line with a small intermediate gap
    thinThickThinLargeGap - a thin-thick-thin line with a large gap
    thinThickThinMediumGap - a thin-thick-thin line with a medium gap
    thinThickThinSmallGap - a thin-thick-thin line with a small gap
    threeDEmboss - a three-staged gradient line, getting darker towards the paragraph
    threeDEngrave - a three-staged gradient like, getting darker away from the paragraph
    triple - a triple line
    wave - a wavy line
*/

END zt_word;
/
