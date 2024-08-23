CREATE OR REPLACE PACKAGE BODY zt_word AS

    --global variables 
    grDoc t_documents := t_documents();
    gcClob clob;
    c_local_file_header constant raw(4) := hextoraw( '504B0304' );
    c_end_of_central_directory constant raw(4) := hextoraw( '504B0506' );

    grDefaultPage r_page := f_get_page(
        p_width => 11906,
        p_height => 16838,
        p_margin_top => 1417,
        p_margin_bottom => 1417,
        p_margin_left => 1417,
        p_margin_right => 1417,
        p_header_height => 708,
        p_footer_height => 708,
        p_orientation => 'portrait');



FUNCTION f_unit_convert(
    p_doc_id number default null,
    p_usage varchar2 default 'page',  --values: page, image, image_rotate
    p_value number
    ) RETURN number IS

    lnIndex number;

BEGIN
    if p_doc_id is null then
        RETURN p_value;
    end if;
    
    CASE
        WHEN p_usage = 'page' and grDoc(p_doc_id).unit = 'cm' THEN lnIndex := 567;
        WHEN p_usage = 'page' and grDoc(p_doc_id).unit = 'mm' THEN lnIndex := 56.7;
        WHEN p_usage = 'image' and grDoc(p_doc_id).unit = 'cm' THEN lnIndex := 360000;
        WHEN p_usage = 'image' and grDoc(p_doc_id).unit = 'mm' THEN lnIndex := 36000;
        WHEN p_usage = 'image_rotate' THEN lnIndex := 60000;
        ELSE lnIndex := 1;
    END CASE;
        
    RETURN round(p_value * lnIndex);
END;


FUNCTION f_explode(p_text in varchar2,
                   p_delimiter in varchar2) RETURN t_table_vc2 IS

    lrList t_table_vc2 := t_table_vc2();
    lnCounter pls_integer := 0;
    lcText varchar2(32000) := p_text;
  
BEGIN
    LOOP
        lnCounter := instr(lcText, p_delimiter);

        if lnCounter > 0 then
            lrList.extend(1);
            lrList(lrList.count) := substr(lcText, 1, lnCounter - 1);
            lcText := substr(lcText, lnCounter + length(p_delimiter));
        else
            lrList.extend(1);
            lrList(lrList.count) := lcText;
            return lrList;
        end if;
        
    END LOOP;

END f_explode;


FUNCTION c2b(
    p_clob clob,
    p_encoding IN NUMBER default 0) RETURN blob IS

    v_blob Blob;
    v_in Pls_Integer := 1;
    v_out Pls_Integer := 1;
    v_lang Pls_Integer := 0;
    v_warning Pls_Integer := 0;
    v_id number(10);

BEGIN
    if p_clob is null then
        return null;
    end if;

    v_in:=1;
    v_out:=1;
    dbms_lob.createtemporary(v_blob,TRUE);
    DBMS_LOB.convertToBlob(v_blob,
                           p_clob,
                           DBMS_lob.getlength(p_clob),
                           v_in,
                           v_out,
                           p_encoding,
                           v_lang,
                           v_warning);

    RETURN v_blob;

END c2b;


FUNCTION f_new_container(
    p_doc_id number,
    p_type varchar2 default 'DOCUMENT') RETURN pls_integer IS

    lnID pls_integer;

BEGIN
    grDoc(p_doc_id).containers.extend;
    
    lnID := grDoc(p_doc_id).containers.count;
    
    grDoc(p_doc_id).containers(lnID).container_type := p_type;
    grDoc(p_doc_id).containers(lnID).elements := t_elements();
    
    if p_type in ('HEADER', 'FOOTER') then
        grDoc(p_doc_id).containers(lnID).rel_id := grDoc(p_doc_id).rels_id;
        grDoc(p_doc_id).rels_id := grDoc(p_doc_id).rels_id + 1;
    end if;

    RETURN lnID;
END;
    

FUNCTION f_new_document(
    p_author varchar2 default null,
    p_default_page r_page default null,
    p_unit varchar2 default 'cm',
    p_lang varchar2 default 'en-US') RETURN pls_integer IS
    
    lnID pls_integer;
    lnID2 pls_integer;
    
BEGIN
    grDoc.extend;
    lnID := grDoc.count;
    
    --document properties
    grDoc(lnID).author := p_author;
    grDoc(lnID).create_date := sysdate;
    grDoc(lnID).rels_id := 10;
    
    if p_default_page.width is not null then
        grDoc(lnID).default_page := p_default_page;
    else
        grDoc(lnID).default_page := grDefaultPage;
    end if;

    grDoc(lnID).unit := p_unit;
    grDoc(lnID).lang := p_lang;
    
    --containers init - first one is document
    grDoc(lnID).containers := t_containers();
    lnID2 := f_new_container(lnID);
    
    --lists init
    grDoc(lnID).lists := t_list();
    
    --images init
    grDoc(lnID).images := t_images();
    
    RETURN lnID;
END;


FUNCTION f_new_paragraph(
    p_doc_id number,
    p_container_id pls_integer default null,
    p_alignment_h varchar2 default 'LEFT',
    p_space_before number default 0,
    p_space_after number default 0,
    p_style varchar2 default null,
    p_list_id pls_integer default null,
    p_text varchar2 default null,
    p_font r_font default null,
    p_replace_newline boolean default false,
    p_newline_character varchar2 default chr(10)
) RETURN pls_integer IS
    
    lnID pls_integer;
    lnContainerID pls_integer := nvl(p_container_id, 1);
    
BEGIN
    grDoc(p_doc_id).containers(lnContainerID).elements.extend;
    lnID := grDoc(p_doc_id).containers(lnContainerID).elements.count;
    
    grDoc(p_doc_id).containers(lnContainerID).elements(lnID).element_type := 'PARAGRAPH';
    grDoc(p_doc_id).containers(lnContainerID).elements(lnID).paragraph.alignment_h := p_alignment_h;
    grDoc(p_doc_id).containers(lnContainerID).elements(lnID).paragraph.space_before := p_space_before;
    grDoc(p_doc_id).containers(lnContainerID).elements(lnID).paragraph.space_after := p_space_after;
    grDoc(p_doc_id).containers(lnContainerID).elements(lnID).paragraph.style := p_style;
    grDoc(p_doc_id).containers(lnContainerID).elements(lnID).paragraph.list_id := p_list_id;
    
    if p_text is not null then
        p_add_text(
            p_doc_id => p_doc_id,
            p_container_id => lnContainerID,
            p_paragraph_id => lnID, 
            p_text => p_text, 
            p_font => p_font,
            p_replace_newline => p_replace_newline,
            p_newline_character => p_newline_character
        );
    end if;
    
    RETURN lnID;
END;    


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
    ) RETURN r_image_data IS

    lrImageParams r_image_data;
    
BEGIN
    lrImageParams.image_id := p_image_id;
    lrImageParams.width := p_width;
    lrImageParams.height := p_height;
    lrImageParams.rotate_angle := p_rotate_angle;
    lrImageParams.extent_area_left := p_extent_area_left;
    lrImageParams.extent_area_top := p_extent_area_top;
    lrImageParams.extent_area_right := p_extent_area_right;
    lrImageParams.extent_area_bottom := p_extent_area_bottom;
    lrImageParams.inline_yn := p_inline_yn;
    lrImageParams.relative_from_h := p_relative_from_h;
    lrImageParams.position_type_h := p_position_type_h;
    lrImageParams.position_align_h := p_position_align_h;
    lrImageParams.position_h := p_position_h;
    lrImageParams.relative_from_v := p_relative_from_v;
    lrImageParams.position_type_v := p_position_type_v;
    lrImageParams.position_align_v := p_position_align_v;
    lrImageParams.position_v := p_position_v;
    
    RETURN lrImageParams;
END f_image_data;

FUNCTION f_add_image_to_document(
    p_doc_id number,
    p_filename varchar2,
    p_image blob
    ) RETURN pls_integer IS

    lnID pls_integer;
    
BEGIN
    grDoc(p_doc_id).images.extend;
    lnID := grDoc(p_doc_id).images.count;
    
    grDoc(p_doc_id).images(lnID).image_name := p_filename;
    grDoc(p_doc_id).images(lnID).image_file := p_image;

    grDoc(p_doc_id).images(lnID).rel_id := grDoc(p_doc_id).rels_id;
    grDoc(p_doc_id).rels_id := grDoc(p_doc_id).rels_id + 1;

    RETURN lnID;
END f_add_image_to_document;



FUNCTION f_new_image_instance(
    p_doc_id number,
    p_container_id pls_integer default null,
    p_image_data r_image_data
    ) RETURN pls_integer IS
    
    lnID pls_integer;
    lnContainerID pls_integer := nvl(p_container_id, 1);
    
BEGIN
    grDoc(p_doc_id).containers(lnContainerID).elements.extend;
    lnID := grDoc(p_doc_id).containers(lnContainerID).elements.count;
    
    grDoc(p_doc_id).containers(lnContainerID).elements(lnID).element_type := 'IMAGE';
    grDoc(p_doc_id).containers(lnContainerID).elements(lnID).image_data := p_image_data;
    
    RETURN lnID;
END f_new_image_instance;



FUNCTION f_border(
    p_border_type varchar2 default null,
    p_width number default null,
    p_color varchar2 default null
    ) RETURN r_border IS
    
    lrBorder r_border;
    
BEGIN
    lrBorder.border_type := p_border_type;
    lrBorder.width := p_width;
    lrBorder.color := p_color;
    
    RETURN lrBorder;
END;


PROCEDURE p_table_cell(
    p_doc_id number,
    p_table_id pls_integer,
    p_row pls_integer,
    p_column pls_integer,
    p_alignment_h varchar2 default 'LEFT',
    p_alignment_v varchar2 default 'TOP',
    p_border_top r_border default null,
    p_border_bottom r_border default null,
    p_border_left r_border default null,
    p_border_right r_border default null,
    p_background_color varchar2 default null,
    p_text varchar2 default null,
    p_font r_font default null,
    p_image_data r_image_data default null,
    p_container_id pls_integer default null
    ) IS
    
    lnID pls_integer;
    lnContainerID pls_integer := nvl(p_container_id, 1);
    
BEGIN
    --texts init
    if grDoc(p_doc_id).containers(lnContainerID).elements(p_table_id).table_data.cells(p_row || ',' || p_column).paragraph.texts is null then
        grDoc(p_doc_id).containers(lnContainerID).elements(p_table_id).table_data.cells(p_row || ',' || p_column).paragraph.texts := t_text();
    end if;

    --alignment
    grDoc(p_doc_id).containers(lnContainerID).elements(p_table_id).table_data.cells(p_row || ',' || p_column).paragraph.alignment_h := p_alignment_h;
    grDoc(p_doc_id).containers(lnContainerID).elements(p_table_id).table_data.cells(p_row || ',' || p_column).alignment_v := p_alignment_v;

    --background color
    grDoc(p_doc_id).containers(lnContainerID).elements(p_table_id).table_data.cells(p_row || ',' || p_column).background_color := p_background_color;

    --borders
    grDoc(p_doc_id).containers(lnContainerID).elements(p_table_id).table_data.cells(p_row || ',' || p_column).border_top := p_border_top;
    grDoc(p_doc_id).containers(lnContainerID).elements(p_table_id).table_data.cells(p_row || ',' || p_column).border_bottom := p_border_bottom;
    grDoc(p_doc_id).containers(lnContainerID).elements(p_table_id).table_data.cells(p_row || ',' || p_column).border_left := p_border_left;
    grDoc(p_doc_id).containers(lnContainerID).elements(p_table_id).table_data.cells(p_row || ',' || p_column).border_right := p_border_right;
    
    --text
    if p_text is not null or p_image_data.image_id is not null then
        grDoc(p_doc_id).containers(lnContainerID).elements(p_table_id).table_data.cells(p_row || ',' || p_column).paragraph.texts.extend;
        lnID := grDoc(p_doc_id).containers(lnContainerID).elements(p_table_id).table_data.cells(p_row || ',' || p_column).paragraph.texts.count;
        
        grDoc(p_doc_id).containers(lnContainerID).elements(p_table_id).table_data.cells(p_row || ',' || p_column).paragraph.texts(lnID).text := p_text;
        grDoc(p_doc_id).containers(lnContainerID).elements(p_table_id).table_data.cells(p_row || ',' || p_column).paragraph.texts(lnID).font := p_font;

        grDoc(p_doc_id).containers(lnContainerID).elements(p_table_id).table_data.cells(p_row || ',' || p_column).paragraph.texts(lnID).image_data := p_image_data;
    end if;
END p_table_cell;


PROCEDURE p_table_column_width(
    p_doc_id pls_integer,
    p_table_id pls_integer,
    p_width varchar2,
    p_container_id pls_integer default null) IS
    
    lcSirina varchar2(10000) := replace(p_width, ' ', null);
    lrVrednosti t_table_vc2;
    lnContainerID pls_integer := nvl(p_container_id, 1);
    
BEGIN
    grDoc(p_doc_id).containers(lnContainerID).elements(p_table_id).table_data.column_width := t_table_number();

    lrVrednosti := f_explode(lcSirina, ',');
    
    FOR t IN 1 .. lrVrednosti.count LOOP
        grDoc(p_doc_id).containers(lnContainerID).elements(p_table_id).table_data.column_width.extend;
        grDoc(p_doc_id).containers(lnContainerID).elements(p_table_id).table_data.column_width(t) := to_number(lrVrednosti(t));
    END LOOP;

END;


PROCEDURE p_table_merge_cells(
    p_doc_id number,
    p_table_id pls_integer,
    p_from_row pls_integer,
    p_from_column pls_integer,
    p_to_row pls_integer,
    p_to_column pls_integer,
    p_container_id pls_integer default null) IS

    lnContainerID pls_integer := nvl(p_container_id, 1);
    
BEGIN
    --if vertical merge exists -> mark cells
    if p_from_row <> p_to_row then
        --mark first row
        grDoc(p_doc_id).containers(lnContainerID).elements(p_table_id).table_data.cells(p_from_row || ',' || p_from_column).merge_v := '<w:vMerge w:val="restart"/>';
        
        --mark other rows till last row
        FOR t IN (p_from_row + 1) .. p_to_row LOOP
            grDoc(p_doc_id).containers(lnContainerID).elements(p_table_id).table_data.cells(t || ',' || p_from_column).merge_v := '<w:vMerge/>';
        END LOOP;
    end if;
    
    --if horizontal merge exists -> mark cells
    if p_from_column <> p_to_column then
        --for each row merge cells -> mark gridSpan number; other columns mark with -1 (ignore in XML document)
        FOR t IN p_from_row .. p_to_row LOOP
            grDoc(p_doc_id).containers(lnContainerID).elements(p_table_id).table_data.cells(t || ',' || p_from_column).merge_h := p_to_column - p_from_column + 1;
            FOR p IN (p_from_column + 1) .. p_to_column LOOP
                grDoc(p_doc_id).containers(lnContainerID).elements(p_table_id).table_data.cells(t || ',' || p).merge_h := -1;
            END LOOP;
        END LOOP;
    end if;
END;

FUNCTION f_new_table(
    p_doc_id number,
    p_rows pls_integer,
    p_columns pls_integer,
    p_table_width pls_integer default null,
    p_columns_width varchar2 default null,
    p_border_top r_border default null,
    p_border_bottom r_border default null,
    p_border_left r_border default null,
    p_border_right r_border default null,
    p_border_inside_h r_border default null,
    p_border_inside_v r_border default null,
    p_container_id pls_integer default null
    ) RETURN pls_integer IS

    lnID pls_integer;
    lnContainerID pls_integer := nvl(p_container_id, 1);
    
BEGIN
    grDoc(p_doc_id).containers(lnContainerID).elements.extend;
    lnID := grDoc(p_doc_id).containers(lnContainerID).elements.count;

    grDoc(p_doc_id).containers(lnContainerID).elements(lnID).element_type := 'TABLE';
    grDoc(p_doc_id).containers(lnContainerID).elements(lnID).table_data.rows_num := p_rows;
    grDoc(p_doc_id).containers(lnContainerID).elements(lnID).table_data.columns_num := p_columns;

    grDoc(p_doc_id).containers(lnContainerID).elements(lnID).table_data.width := p_table_width;

    grDoc(p_doc_id).containers(lnContainerID).elements(lnID).table_data.border_top := p_border_top;
    grDoc(p_doc_id).containers(lnContainerID).elements(lnID).table_data.border_bottom := p_border_bottom;
    grDoc(p_doc_id).containers(lnContainerID).elements(lnID).table_data.border_left := p_border_left;
    grDoc(p_doc_id).containers(lnContainerID).elements(lnID).table_data.border_right := p_border_right;
    grDoc(p_doc_id).containers(lnContainerID).elements(lnID).table_data.border_inside_h := p_border_inside_h;
    grDoc(p_doc_id).containers(lnContainerID).elements(lnID).table_data.border_inside_v := p_border_inside_v;
    
    p_table_column_width(
        p_doc_id => p_doc_id, 
        p_container_id => lnContainerID,
        p_table_id => lnID, 
        p_width => p_columns_width 
    );
    
    --create cells
    FOR v IN 1 .. p_rows LOOP
        FOR s IN 1 .. p_columns LOOP
            grDoc(p_doc_id).containers(lnContainerID).elements(lnID).table_data.cells(v || ',' || s) := null;
        END LOOP;
    END LOOP;

    RETURN lnID;
END;



FUNCTION f_new_numbering(
    p_doc_id number,
    p_start_value varchar2 default '1') RETURN pls_integer IS

    lnID pls_integer;

BEGIN
    grDoc(p_doc_id).lists.extend;
    lnID := grDoc(p_doc_id).lists.count;

    grDoc(p_doc_id).lists(lnID).list_type := 'decimal';
    grDoc(p_doc_id).lists(lnID).num_start_value := p_start_value;

    RETURN lnID;
END;


FUNCTION f_new_bullet(
    p_doc_id number,
    p_char varchar2 default 'o',
    p_font varchar2 default 'Courier New') RETURN pls_integer IS

    lnID pls_integer;

BEGIN
    grDoc(p_doc_id).lists.extend;
    lnID := grDoc(p_doc_id).lists.count;

    grDoc(p_doc_id).lists(lnID).list_type := 'bullet';
    grDoc(p_doc_id).lists(lnID).bullet_char := p_char;
    grDoc(p_doc_id).lists(lnID).bullet_font := p_font;

    RETURN lnID;
END;


FUNCTION f_new_page_break(p_doc_id number) RETURN pls_integer IS

    lnID pls_integer;
    
BEGIN
    grDoc(p_doc_id).containers(1).elements.extend;
    lnID := grDoc(p_doc_id).containers(1).elements.count;

    grDoc(p_doc_id).containers(1).elements(lnID).element_type := 'BREAK';

    grDoc(p_doc_id).containers(1).elements(lnID).break_data.break_type := 'PAGE';
    
    RETURN lnID;
END;


FUNCTION f_new_section_break(
    p_doc_id number,
    p_section_type varchar2,
    p_page_template varchar2 default null,
    p_page r_page default null) RETURN pls_integer IS

    lnID pls_integer;
    
BEGIN
    grDoc(p_doc_id).containers(1).elements.extend;
    lnID := grDoc(p_doc_id).containers(1).elements.count;

    grDoc(p_doc_id).containers(1).elements(lnID).element_type := 'BREAK';

    grDoc(p_doc_id).containers(1).elements(lnID).break_data.break_type := 'SECTION';
    grDoc(p_doc_id).containers(1).elements(lnID).break_data.section_type := p_section_type;
    
    if nvl(p_page_template, 'x') = 'default' then
        grDoc(p_doc_id).containers(1).elements(lnID).break_data.page := grDoc(p_doc_id).default_page;
    else 
        grDoc(p_doc_id).containers(1).elements(lnID).break_data.page := p_page;
    end if;
    
    RETURN lnID;
END;


PROCEDURE p_set_default_page(
    p_doc_id number,
    p_page r_page) IS
BEGIN
    grDoc(p_doc_id).default_page := p_page;
END;

FUNCTION f_get_default_page(
    p_doc_id number) RETURN r_page IS
BEGIN
    RETURN grDoc(p_doc_id).default_page;
END;

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
    p_orientation varchar2) RETURN r_page IS
    
    lrPage zt_word.r_page;
    
BEGIN
    lrPage.width := f_unit_convert(p_doc_id, 'page', p_width);
    lrPage.height := f_unit_convert(p_doc_id, 'page', p_height);

    lrPage.margin_top := f_unit_convert(p_doc_id, 'page', p_margin_top);
    lrPage.margin_bottom := f_unit_convert(p_doc_id, 'page', p_margin_bottom);
    lrPage.margin_left := f_unit_convert(p_doc_id, 'page', p_margin_left);
    lrPage.margin_right := f_unit_convert(p_doc_id, 'page', p_margin_right);

    lrPage.header_h := f_unit_convert(p_doc_id, 'page', p_header_height);
    lrPage.footer_h := f_unit_convert(p_doc_id, 'page', p_footer_height);

    lrPage.header_ref := f_unit_convert(p_doc_id, 'page', p_header_ref);
    lrPage.footer_ref := f_unit_convert(p_doc_id, 'page', p_footer_ref);

    lrPage.orientation := p_orientation;
    
    RETURN lrPage;
END;

FUNCTION f_font(
    p_from_paragraph boolean default false,
    p_font_name varchar2 default 'Calibri',
    p_font_size number default 11,
    p_bold boolean default false,
    p_italic boolean default false,
    p_underline boolean default false,
    p_color varchar2 default '000000') RETURN r_font IS

    lrFont r_font;
    
BEGIN
    lrFont.from_paragraph := p_from_paragraph;
    lrFont.font_name := p_font_name;
    lrFont.font_size := p_font_size;
    lrFont.bold := p_bold;
    lrFont.italic := p_italic;
    lrFont.underline := p_underline;
    lrFont.color := p_color;
    
    RETURN lrFont;
END;


PROCEDURE p_add_text(
    p_doc_id number,
    p_container_id pls_integer default null,
    p_paragraph_id number,
    p_text varchar2 default null,
    p_font r_font default null,
    p_image_data r_image_data default null,
    p_replace_newline boolean default false,
    p_newline_character varchar2 default chr(10)
) IS
    
    lnID number;
    lnContainerID pls_integer := nvl(p_container_id, 1);
    
BEGIN
    if grDoc(p_doc_id).containers(lnContainerID).elements(p_paragraph_id).paragraph.texts is null then
        grDoc(p_doc_id).containers(lnContainerID).elements(p_paragraph_id).paragraph.texts := t_text();
    end if;

    grDoc(p_doc_id).containers(lnContainerID).elements(p_paragraph_id).paragraph.texts.extend;
    lnID := grDoc(p_doc_id).containers(lnContainerID).elements(p_paragraph_id).paragraph.texts.count;

    grDoc(p_doc_id).containers(lnContainerID).elements(p_paragraph_id).paragraph.texts(lnID).text := p_text;
    grDoc(p_doc_id).containers(lnContainerID).elements(p_paragraph_id).paragraph.texts(lnID).font := p_font;
    grDoc(p_doc_id).containers(lnContainerID).elements(p_paragraph_id).paragraph.texts(lnID).replace_newline := p_replace_newline;
    grDoc(p_doc_id).containers(lnContainerID).elements(p_paragraph_id).paragraph.texts(lnID).newline_character := p_newline_character;

    grDoc(p_doc_id).containers(lnContainerID).elements(p_paragraph_id).paragraph.texts(lnID).image_data := p_image_data;
END p_add_text;

PROCEDURE p_add_line_break (
    p_doc_id number,
    p_container_id pls_integer default null,
    p_paragraph_id number
) IS

    lnID number;
    lnContainerID pls_integer := nvl(p_container_id, 1);

BEGIN
    if grDoc(p_doc_id).containers(lnContainerID).elements(p_paragraph_id).paragraph.texts is null then
        grDoc(p_doc_id).containers(lnContainerID).elements(p_paragraph_id).paragraph.texts := t_text();
    end if;

    grDoc(p_doc_id).containers(lnContainerID).elements(p_paragraph_id).paragraph.texts.extend;
    lnID := grDoc(p_doc_id).containers(lnContainerID).elements(p_paragraph_id).paragraph.texts.count;

    grDoc(p_doc_id).containers(lnContainerID).elements(p_paragraph_id).paragraph.texts(lnID).line_break := true;

END p_add_line_break;


  function little_endian( p_big number, p_bytes pls_integer := 4 )
  return raw
  is
  begin
    return utl_raw.substr( utl_raw.cast_from_binary_integer( p_big, utl_raw.little_endian ), 1, p_bytes );
  end;

  function blob2num( p_blob blob, p_len integer, p_pos integer )
  return number
  is
  begin
    return utl_raw.cast_to_binary_integer( dbms_lob.substr( p_blob, p_len, p_pos ), utl_raw.little_endian );
  end;


  procedure add1file
    ( p_zipped_blob in out blob
    , p_name varchar2
    , p_content blob
    )
  is
    t_now date;
    t_blob blob;
    t_len integer;
    t_clen integer;
    t_crc32 raw(4) := hextoraw( '00000000' );
    t_compressed boolean := false;
    t_name raw(32767);
  begin
    t_now := sysdate;
    t_len := nvl( dbms_lob.getlength( p_content ), 0 );
    if t_len > 0
    then 
      t_blob := utl_compress.lz_compress( p_content );
      t_clen := dbms_lob.getlength( t_blob ) - 18;
      t_compressed := t_clen < t_len;
      t_crc32 := dbms_lob.substr( t_blob, 4, t_clen + 11 );       
    end if;
    if not t_compressed
    then 
      t_clen := t_len;
      t_blob := p_content;
    end if;
    if p_zipped_blob is null
    then
      dbms_lob.createtemporary( p_zipped_blob, true );
    end if;
    t_name := utl_i18n.string_to_raw( p_name, 'AL32UTF8' );
    dbms_lob.append( p_zipped_blob
                   , utl_raw.concat( c_LOCAL_FILE_HEADER -- Local file header signature
                                   , hextoraw( '1400' )  -- version 2.0
                                   , case when t_name = utl_i18n.string_to_raw( p_name, 'US8PC437' )
                                       then hextoraw( '0000' ) -- no General purpose bits
                                       else hextoraw( '0008' ) -- set Language encoding flag (EFS)
                                     end 
                                   , case when t_compressed
                                        then hextoraw( '0800' ) -- deflate
                                        else hextoraw( '0000' ) -- stored
                                     end
                                   , little_endian( to_number( to_char( t_now, 'ss' ) ) / 2
                                                  + to_number( to_char( t_now, 'mi' ) ) * 32
                                                  + to_number( to_char( t_now, 'hh24' ) ) * 2048
                                                  , 2
                                                  ) -- File last modification time
                                   , little_endian( to_number( to_char( t_now, 'dd' ) )
                                                  + to_number( to_char( t_now, 'mm' ) ) * 32
                                                  + ( to_number( to_char( t_now, 'yyyy' ) ) - 1980 ) * 512
                                                  , 2
                                                  ) -- File last modification date
                                   , t_crc32 -- CRC-32
                                   , little_endian( t_clen )                      -- compressed size
                                   , little_endian( t_len )                       -- uncompressed size
                                   , little_endian( utl_raw.length( t_name ), 2 ) -- File name length
                                   , hextoraw( '0000' )                           -- Extra field length
                                   , t_name                                       -- File name
                                   )
                   );
    if t_compressed
    then                   
      dbms_lob.copy( p_zipped_blob, t_blob, t_clen, dbms_lob.getlength( p_zipped_blob ) + 1, 11 ); -- compressed content
    elsif t_clen > 0
    then                   
      dbms_lob.copy( p_zipped_blob, t_blob, t_clen, dbms_lob.getlength( p_zipped_blob ) + 1, 1 ); --  content
    end if;
    if dbms_lob.istemporary( t_blob ) = 1
    then      
      dbms_lob.freetemporary( t_blob );
    end if;
  end;
--
  procedure finish_zip( p_zipped_blob in out blob )
  is
    t_cnt pls_integer := 0;
    t_offs integer;
    t_offs_dir_header integer;
    t_offs_end_header integer;
    t_comment raw(32767) := utl_raw.cast_to_raw( 'Implementation by Anton Scheffer' );
  begin
    t_offs_dir_header := dbms_lob.getlength( p_zipped_blob );
    t_offs := 1;
    while dbms_lob.substr( p_zipped_blob, utl_raw.length( c_LOCAL_FILE_HEADER ), t_offs ) = c_LOCAL_FILE_HEADER
    loop
      t_cnt := t_cnt + 1;
      dbms_lob.append( p_zipped_blob
                     , utl_raw.concat( hextoraw( '504B0102' )      -- Central directory file header signature
                                     , hextoraw( '1400' )          -- version 2.0
                                     , dbms_lob.substr( p_zipped_blob, 26, t_offs + 4 )
                                     , hextoraw( '0000' )          -- File comment length
                                     , hextoraw( '0000' )          -- Disk number where file starts
                                     , hextoraw( '0000' )          -- Internal file attributes => 
                                                                   --     0000 binary file
                                                                   --     0100 (ascii)text file
                                     , case
                                         when dbms_lob.substr( p_zipped_blob
                                                             , 1
                                                             , t_offs + 30 + blob2num( p_zipped_blob, 2, t_offs + 26 ) - 1
                                                             ) in ( hextoraw( '2F' ) -- /
                                                                  , hextoraw( '5C' ) -- \
                                                                  )
                                         then hextoraw( '10000000' ) -- a directory/folder
                                         else hextoraw( '2000B681' ) -- a file
                                       end                         -- External file attributes
                                     , little_endian( t_offs - 1 ) -- Relative offset of local file header
                                     , dbms_lob.substr( p_zipped_blob
                                                      , blob2num( p_zipped_blob, 2, t_offs + 26 )
                                                      , t_offs + 30
                                                      )            -- File name
                                     )
                     );
      t_offs := t_offs + 30 + blob2num( p_zipped_blob, 4, t_offs + 18 )  -- compressed size
                            + blob2num( p_zipped_blob, 2, t_offs + 26 )  -- File name length 
                            + blob2num( p_zipped_blob, 2, t_offs + 28 ); -- Extra field length
    end loop;
    t_offs_end_header := dbms_lob.getlength( p_zipped_blob );
    dbms_lob.append( p_zipped_blob
                   , utl_raw.concat( c_END_OF_CENTRAL_DIRECTORY                                -- End of central directory signature
                                   , hextoraw( '0000' )                                        -- Number of this disk
                                   , hextoraw( '0000' )                                        -- Disk where central directory starts
                                   , little_endian( t_cnt, 2 )                                 -- Number of central directory records on this disk
                                   , little_endian( t_cnt, 2 )                                 -- Total number of central directory records
                                   , little_endian( t_offs_end_header - t_offs_dir_header )    -- Size of central directory
                                   , little_endian( t_offs_dir_header )                        -- Offset of start of central directory, relative to start of archive
                                   , little_endian( nvl( utl_raw.length( t_comment ), 0 ), 2 ) -- ZIP file comment length
                                   , t_comment
                                   )
                   );
  end finish_zip;




PROCEDURE p_add_document_to_zip(
    p_zip IN OUT blob,
    p_name varchar2,
    p_document clob
    ) IS

    lbBlob blob;

BEGIN
    lbBlob := c2b(p_document);
    add1file(p_zip, p_name, lbBlob);
END p_add_document_to_zip;


PROCEDURE p_add_document_to_zip(
    p_zip IN OUT blob,
    p_name varchar2,
    p_document blob
    ) IS

BEGIN
    add1file(p_zip, p_name, p_document);
END p_add_document_to_zip;





FUNCTION f_content_types(p_doc_id number) RETURN clob IS
    lcClob clob;

BEGIN
    lcClob := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="png" ContentType="image/png"/>
    <Default Extension="jpg" ContentType="image/jpg"/>
	<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
	<Default Extension="xml" ContentType="application/xml"/>
	<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' ||
	(CASE WHEN grDoc(p_doc_id).lists.count = 0 THEN null ELSE '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' END) ||
	'<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
	<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
	<Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>
	<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
	<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
	<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
	<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
';

    FOR t IN grDoc(p_doc_id).containers.first .. grDoc(p_doc_id).containers.last LOOP
        if grDoc(p_doc_id).containers(t).container_type = 'HEADER' then
            lcClob := lcClob || chr(10) || '<Override PartName="/word/header' || grDoc(p_doc_id).containers(t).rel_id || '.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>';
        elsif grDoc(p_doc_id).containers(t).container_type = 'FOOTER' then
            lcClob := lcClob || chr(10) || '<Override PartName="/word/footer' || grDoc(p_doc_id).containers(t).rel_id || '.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>';
        end if;
    END LOOP;

    lcClob := lcClob || chr(10) || '</Types>';

    RETURN lcClob;
END f_content_types;


FUNCTION f_rels RETURN clob IS
BEGIN
    RETURN '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
	<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
	<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
	<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>';
END f_rels;


FUNCTION f_app RETURN clob IS
BEGIN
    RETURN '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
	<Template>Normal.dotm</Template>
	<TotalTime>1</TotalTime>
	<Pages>1</Pages>
	<Words>10</Words>
	<Characters>58</Characters>
	<Application>Microsoft Office Word</Application>
	<DocSecurity>0</DocSecurity>
	<Lines>1</Lines>
	<Paragraphs>1</Paragraphs>
	<ScaleCrop>false</ScaleCrop>
	<Company/>
	<LinksUpToDate>false</LinksUpToDate>
	<CharactersWithSpaces>67</CharactersWithSpaces>
	<SharedDoc>false</SharedDoc>
	<HyperlinksChanged>false</HyperlinksChanged>
	<AppVersion>15.0000</AppVersion>
</Properties>';
END f_app;

FUNCTION f_core(p_doc_id number) RETURN clob IS
BEGIN
    RETURN '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
	<dc:title/>
	<dc:subject/>
	<dc:creator>' || grDoc(p_doc_id).author || '</dc:creator>
	<cp:keywords/>
	<dc:description/>
	<cp:lastModifiedBy>' || grDoc(p_doc_id).author || '</cp:lastModifiedBy>
	<cp:revision>2</cp:revision>
	<dcterms:created xsi:type="dcterms:W3CDTF">' || to_char(grDoc(p_doc_id).create_date, 'yyyy-mm-dd') || 'T' || to_char(grDoc(p_doc_id).create_date, 'hh24:mi:ss') || 'Z</dcterms:created>
	<dcterms:modified xsi:type="dcterms:W3CDTF">' || to_char(grDoc(p_doc_id).create_date, 'yyyy-mm-dd') || 'T' || to_char(grDoc(p_doc_id).create_date, 'hh24:mi:ss') || 'Z</dcterms:modified>
</cp:coreProperties>';
END f_core;



FUNCTION f_document_xml_rels(
    p_doc_id pls_integer
    ) RETURN clob IS
    
    lcRels clob;
    
BEGIN
    lcRels := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
	<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/>
	<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
	<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
	<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
	<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>' ||
	(CASE WHEN grDoc(p_doc_id).lists.count = 0 THEN null ELSE chr(10) || '	<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>' END);
	
	--references for headers and footers
	FOR t IN grDoc(p_doc_id).containers.first .. grDoc(p_doc_id).containers.last LOOP
        if grDoc(p_doc_id).containers(t).container_type = 'HEADER' then
            lcRels := lcRels || chr(10) ||
                '	<Relationship Id="rId' || 
                grDoc(p_doc_id).containers(t).rel_id || 
                '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header' || 
                grDoc(p_doc_id).containers(t).rel_id || 
                '.xml"/>';
        elsif grDoc(p_doc_id).containers(t).container_type = 'FOOTER' then
            lcRels := lcRels || chr(10) ||
                '	<Relationship Id="rId' || 
                grDoc(p_doc_id).containers(t).rel_id || 
                '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer' || 
                grDoc(p_doc_id).containers(t).rel_id || 
                '.xml"/>';
        end if;
	END LOOP;

    --references for images
    FOR t IN 1 .. grDoc(p_doc_id).images.count LOOP
        lcRels := lcRels || chr(10) ||
            chr(9) || '<Relationship Id="rId' || 
            grDoc(p_doc_id).images(t).rel_id || 
            '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/' ||
            grDoc(p_doc_id).images(t).image_name ||
            '"/>';
    END LOOP;
    
	
    lcRels := lcRels || chr(10) || '</Relationships>';

    RETURN lcRels;
END f_document_xml_rels;


FUNCTION f_container_xml_rels(
    p_doc_id pls_integer
    ) RETURN clob IS
    
    lcRels clob;
    
BEGIN
    lcRels := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
	
    --references for images
    FOR t IN 1 .. grDoc(p_doc_id).images.count LOOP
        lcRels := lcRels || chr(10) ||
            chr(9) || '<Relationship Id="rId' || 
            grDoc(p_doc_id).images(t).rel_id || 
            '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/' ||
            grDoc(p_doc_id).images(t).image_name ||
            '"/>';
    END LOOP;
	
    lcRels := lcRels || chr(10) || '</Relationships>';

    RETURN lcRels;
END f_container_xml_rels;


FUNCTION f_theme RETURN clob IS
BEGIN
    RETURN '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
	<a:themeElements>
		<a:clrScheme name="Office">
			<a:dk1>
				<a:sysClr val="windowText" lastClr="000000"/>
			</a:dk1>
			<a:lt1>
				<a:sysClr val="window" lastClr="FFFFFF"/>
			</a:lt1>
			<a:dk2>
				<a:srgbClr val="44546A"/>
			</a:dk2>
			<a:lt2>
				<a:srgbClr val="E7E6E6"/>
			</a:lt2>
			<a:accent1>
				<a:srgbClr val="5B9BD5"/>
			</a:accent1>
			<a:accent2>
				<a:srgbClr val="ED7D31"/>
			</a:accent2>
			<a:accent3>
				<a:srgbClr val="A5A5A5"/>
			</a:accent3>
			<a:accent4>
				<a:srgbClr val="FFC000"/>
			</a:accent4>
			<a:accent5>
				<a:srgbClr val="4472C4"/>
			</a:accent5>
			<a:accent6>
				<a:srgbClr val="70AD47"/>
			</a:accent6>
			<a:hlink>
				<a:srgbClr val="0563C1"/>
			</a:hlink>
			<a:folHlink>
				<a:srgbClr val="954F72"/>
			</a:folHlink>
		</a:clrScheme>
		<a:fontScheme name="Office">
			<a:majorFont>
				<a:latin typeface="Calibri Light" panose="020F0302020204030204"/>
				<a:ea typeface=""/>
				<a:cs typeface=""/>
				<a:font script="Jpan" typeface="MS ????"/>
				<a:font script="Hang" typeface="?? ??"/>
				<a:font script="Hans" typeface="??"/>
				<a:font script="Hant" typeface="????"/>
				<a:font script="Arab" typeface="Times New Roman"/>
				<a:font script="Hebr" typeface="Times New Roman"/>
				<a:font script="Thai" typeface="Angsana New"/>
				<a:font script="Ethi" typeface="Nyala"/>
				<a:font script="Beng" typeface="Vrinda"/>
				<a:font script="Gujr" typeface="Shruti"/>
				<a:font script="Khmr" typeface="MoolBoran"/>
				<a:font script="Knda" typeface="Tunga"/>
				<a:font script="Guru" typeface="Raavi"/>
				<a:font script="Cans" typeface="Euphemia"/>
				<a:font script="Cher" typeface="Plantagenet Cherokee"/>
				<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
				<a:font script="Tibt" typeface="Microsoft Himalaya"/>
				<a:font script="Thaa" typeface="MV Boli"/>
				<a:font script="Deva" typeface="Mangal"/>
				<a:font script="Telu" typeface="Gautami"/>
				<a:font script="Taml" typeface="Latha"/>
				<a:font script="Syrc" typeface="Estrangelo Edessa"/>
				<a:font script="Orya" typeface="Kalinga"/>
				<a:font script="Mlym" typeface="Kartika"/>
				<a:font script="Laoo" typeface="DokChampa"/>
				<a:font script="Sinh" typeface="Iskoola Pota"/>
				<a:font script="Mong" typeface="Mongolian Baiti"/>
				<a:font script="Viet" typeface="Times New Roman"/>
				<a:font script="Uigh" typeface="Microsoft Uighur"/>
				<a:font script="Geor" typeface="Sylfaen"/>
			</a:majorFont>
			<a:minorFont>
				<a:latin typeface="Calibri" panose="020F0502020204030204"/>
				<a:ea typeface=""/>
				<a:cs typeface=""/>
				<a:font script="Jpan" typeface="MS ??"/>
				<a:font script="Hang" typeface="?? ??"/>
				<a:font script="Hans" typeface="??"/>
				<a:font script="Hant" typeface="????"/>
				<a:font script="Arab" typeface="Arial"/>
				<a:font script="Hebr" typeface="Arial"/>
				<a:font script="Thai" typeface="Cordia New"/>
				<a:font script="Ethi" typeface="Nyala"/>
				<a:font script="Beng" typeface="Vrinda"/>
				<a:font script="Gujr" typeface="Shruti"/>
				<a:font script="Khmr" typeface="DaunPenh"/>
				<a:font script="Knda" typeface="Tunga"/>
				<a:font script="Guru" typeface="Raavi"/>
				<a:font script="Cans" typeface="Euphemia"/>
				<a:font script="Cher" typeface="Plantagenet Cherokee"/>
				<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
				<a:font script="Tibt" typeface="Microsoft Himalaya"/>
				<a:font script="Thaa" typeface="MV Boli"/>
				<a:font script="Deva" typeface="Mangal"/>
				<a:font script="Telu" typeface="Gautami"/>
				<a:font script="Taml" typeface="Latha"/>
				<a:font script="Syrc" typeface="Estrangelo Edessa"/>
				<a:font script="Orya" typeface="Kalinga"/>
				<a:font script="Mlym" typeface="Kartika"/>
				<a:font script="Laoo" typeface="DokChampa"/>
				<a:font script="Sinh" typeface="Iskoola Pota"/>
				<a:font script="Mong" typeface="Mongolian Baiti"/>
				<a:font script="Viet" typeface="Arial"/>
				<a:font script="Uigh" typeface="Microsoft Uighur"/>
				<a:font script="Geor" typeface="Sylfaen"/>
			</a:minorFont>
		</a:fontScheme>
		<a:fmtScheme name="Office">
			<a:fillStyleLst>
				<a:solidFill>
					<a:schemeClr val="phClr"/>
				</a:solidFill>
				<a:gradFill rotWithShape="1">
					<a:gsLst>
						<a:gs pos="0">
							<a:schemeClr val="phClr">
								<a:lumMod val="110000"/>
								<a:satMod val="105000"/>
								<a:tint val="67000"/>
							</a:schemeClr>
						</a:gs>
						<a:gs pos="50000">
							<a:schemeClr val="phClr">
								<a:lumMod val="105000"/>
								<a:satMod val="103000"/>
								<a:tint val="73000"/>
							</a:schemeClr>
						</a:gs>
						<a:gs pos="100000">
							<a:schemeClr val="phClr">
								<a:lumMod val="105000"/>
								<a:satMod val="109000"/>
								<a:tint val="81000"/>
							</a:schemeClr>
						</a:gs>
					</a:gsLst>
					<a:lin ang="5400000" scaled="0"/>
				</a:gradFill>
				<a:gradFill rotWithShape="1">
					<a:gsLst>
						<a:gs pos="0">
							<a:schemeClr val="phClr">
								<a:satMod val="103000"/>
								<a:lumMod val="102000"/>
								<a:tint val="94000"/>
							</a:schemeClr>
						</a:gs>
						<a:gs pos="50000">
							<a:schemeClr val="phClr">
								<a:satMod val="110000"/>
								<a:lumMod val="100000"/>
								<a:shade val="100000"/>
							</a:schemeClr>
						</a:gs>
						<a:gs pos="100000">
							<a:schemeClr val="phClr">
								<a:lumMod val="99000"/>
								<a:satMod val="120000"/>
								<a:shade val="78000"/>
							</a:schemeClr>
						</a:gs>
					</a:gsLst>
					<a:lin ang="5400000" scaled="0"/>
				</a:gradFill>
			</a:fillStyleLst>
			<a:lnStyleLst>
				<a:ln w="6350" cap="flat" cmpd="sng" algn="ctr">
					<a:solidFill>
						<a:schemeClr val="phClr"/>
					</a:solidFill>
					<a:prstDash val="solid"/>
					<a:miter lim="800000"/>
				</a:ln>
				<a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">
					<a:solidFill>
						<a:schemeClr val="phClr"/>
					</a:solidFill>
					<a:prstDash val="solid"/>
					<a:miter lim="800000"/>
				</a:ln>
				<a:ln w="19050" cap="flat" cmpd="sng" algn="ctr">
					<a:solidFill>
						<a:schemeClr val="phClr"/>
					</a:solidFill>
					<a:prstDash val="solid"/>
					<a:miter lim="800000"/>
				</a:ln>
			</a:lnStyleLst>
			<a:effectStyleLst>
				<a:effectStyle>
					<a:effectLst/>
				</a:effectStyle>
				<a:effectStyle>
					<a:effectLst/>
				</a:effectStyle>
				<a:effectStyle>
					<a:effectLst>
						<a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">
							<a:srgbClr val="000000">
								<a:alpha val="63000"/>
							</a:srgbClr>
						</a:outerShdw>
					</a:effectLst>
				</a:effectStyle>
			</a:effectStyleLst>
			<a:bgFillStyleLst>
				<a:solidFill>
					<a:schemeClr val="phClr"/>
				</a:solidFill>
				<a:solidFill>
					<a:schemeClr val="phClr">
						<a:tint val="95000"/>
						<a:satMod val="170000"/>
					</a:schemeClr>
				</a:solidFill>
				<a:gradFill rotWithShape="1">
					<a:gsLst>
						<a:gs pos="0">
							<a:schemeClr val="phClr">
								<a:tint val="93000"/>
								<a:satMod val="150000"/>
								<a:shade val="98000"/>
								<a:lumMod val="102000"/>
							</a:schemeClr>
						</a:gs>
						<a:gs pos="50000">
							<a:schemeClr val="phClr">
								<a:tint val="98000"/>
								<a:satMod val="130000"/>
								<a:shade val="90000"/>
								<a:lumMod val="103000"/>
							</a:schemeClr>
						</a:gs>
						<a:gs pos="100000">
							<a:schemeClr val="phClr">
								<a:shade val="63000"/>
								<a:satMod val="120000"/>
							</a:schemeClr>
						</a:gs>
					</a:gsLst>
					<a:lin ang="5400000" scaled="0"/>
				</a:gradFill>
			</a:bgFillStyleLst>
		</a:fmtScheme>
	</a:themeElements>
	<a:objectDefaults/>
	<a:extraClrSchemeLst/>
	<a:extLst>
		<a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}">
			<thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/>
		</a:ext>
	</a:extLst>
</a:theme>';
END f_theme;



FUNCTION f_font_table RETURN clob IS
BEGIN
    RETURN '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" mc:Ignorable="w14 w15">
	<w:font w:name="Calibri">
		<w:panose1 w:val="020F0502020204030204"/>
		<w:charset w:val="EE"/>
		<w:family w:val="swiss"/>
		<w:pitch w:val="variable"/>
		<w:sig w:usb0="E00002FF" w:usb1="4000ACFF" w:usb2="00000001" w:usb3="00000000" w:csb0="0000019F" w:csb1="00000000"/>
	</w:font>
	<w:font w:name="Times New Roman">
		<w:panose1 w:val="02020603050405020304"/>
		<w:charset w:val="EE"/>
		<w:family w:val="roman"/>
		<w:pitch w:val="variable"/>
		<w:sig w:usb0="E0002AFF" w:usb1="C0007841" w:usb2="00000009" w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/>
	</w:font>
	<w:font w:name="Calibri Light">
		<w:panose1 w:val="020F0302020204030204"/>
		<w:charset w:val="EE"/>
		<w:family w:val="swiss"/>
		<w:pitch w:val="variable"/>
		<w:sig w:usb0="A00002EF" w:usb1="4000207B" w:usb2="00000000" w:usb3="00000000" w:csb0="0000019F" w:csb1="00000000"/>
	</w:font>
</w:fonts>';
END f_font_table;

/*
FUNCTION f_add_image (p_image_name varchar2, p_params r_image_params) RETURN CLOB IS
BEGIN

RETURN '<w:p w:rsidR="009B1A39" w:rsidRDefault="0051609D">
          <w:bookmarkStart w:id="0" w:name="_GoBack"/>
          <w:r>
            <w:rPr>
              <w:noProof/>
            </w:rPr>
            <w:drawing>
              <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">
                <wp:simplePos x="'||p_params.offset_h||'" y="'||p_params.offset_v||'"/>
                <wp:positionH relativeFrom="'||p_params.relative_from_h||'">
                  <wp:posOffset>'||p_params.offset_h||'</wp:posOffset>
                </wp:positionH>
                <wp:positionV relativeFrom="'||p_params.relative_from_v||'">
                  <wp:posOffset>'||p_params.offset_v||'</wp:posOffset>
                </wp:positionV>
                <wp:extent cx="'||p_params.size_x||'" cy="'||p_params.size_y||'"/>
                <wp:effectExtent l="'||p_params.effect_extent_l||'" t="'||p_params.effect_extent_t||'" r="'||p_params.effect_extent_r||'" b="'||p_params.effect_extent_b||'"/>
                <wp:wrapNone/>
                <wp:docPr id="2" name="Picture 2"/>
                <wp:cNvGraphicFramePr>
                  <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
                </wp:cNvGraphicFramePr>
                <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                    <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                      <pic:nvPicPr>
                        <pic:cNvPr id="2" name="D3RoS 1920x1080.jpg"/>
                        <pic:cNvPicPr/>
                      </pic:nvPicPr>
                      <pic:blipFill>
                        <a:blip r:embed="rId1'||p_image_name||'" cstate="print">
                          <a:extLst>
                            <a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
                              <a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>
                            </a:ext>
                          </a:extLst>
                        </a:blip>
                        <a:stretch>
                          <a:fillRect/>
                        </a:stretch>
                      </pic:blipFill>
                      <pic:spPr>
                        <a:xfrm>
                          <a:off x="0" y="0"/>
                          <a:ext cx="'||p_params.size_x||'" cy="'||p_params.size_y||'"/>
                        </a:xfrm>
                        <a:prstGeom prst="rect">
                          <a:avLst/>
                        </a:prstGeom>
                      </pic:spPr>
                    </pic:pic>
                  </a:graphicData>
                </a:graphic>
                <wp14:sizeRelH relativeFrom="margin">
                  <wp14:pctWidth>0</wp14:pctWidth>
                </wp14:sizeRelH>
                <wp14:sizeRelV relativeFrom="margin">
                  <wp14:pctHeight>0</wp14:pctHeight>
                </wp14:sizeRelV>
              </wp:anchor>
            </w:drawing>
          </w:r>
          <w:bookmarkEnd w:id="0"/>
        </w:p>';
END f_add_image; 
*/

FUNCTION f_settings(p_doc_id number) RETURN clob IS
BEGIN
    RETURN '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main" mc:Ignorable="w14 w15">
	<w:zoom w:percent="100"/>
	<w:defaultTabStop w:val="708"/>
	<w:hyphenationZone w:val="425"/>
	<w:characterSpacingControl w:val="doNotCompress"/>
	<w:compat>
		<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
		<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
		<w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
		<w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
		<w:compatSetting w:name="differentiateMultirowTableHeaders" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
	</w:compat>
	<w:rsids>
		<w:rsidRoot w:val="0048321B"/>
		<w:rsid w:val="0030764E"/>
		<w:rsid w:val="0048321B"/>
		<w:rsid w:val="005C7A42"/>
		<w:rsid w:val="00B34BE1"/>
	</w:rsids>
	<m:mathPr>
		<m:mathFont m:val="Cambria Math"/>
		<m:brkBin m:val="before"/>
		<m:brkBinSub m:val="--"/>
		<m:smallFrac m:val="0"/>
		<m:dispDef/>
		<m:lMargin m:val="0"/>
		<m:rMargin m:val="0"/>
		<m:defJc m:val="centerGroup"/>
		<m:wrapIndent m:val="1440"/>
		<m:intLim m:val="subSup"/>
		<m:naryLim m:val="undOvr"/>
	</m:mathPr>
	<w:themeFontLang w:val="' || grDoc(p_doc_id).lang || '"/>
	<w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/>
	<w:shapeDefaults>
		<o:shapedefaults v:ext="edit" spidmax="1026"/>
		<o:shapelayout v:ext="edit">
			<o:idmap v:ext="edit" data="1"/>
		</o:shapelayout>
	</w:shapeDefaults>
	<w:decimalSymbol w:val=","/>
	<w:listSeparator w:val=";"/>
	<w15:chartTrackingRefBased/>
	<w15:docId w15:val="{CA8CD7EC-1908-42B3-9D0A-362D3CB986C6}"/>
</w:settings>';
END f_settings;


FUNCTION f_styles(p_doc_id number) RETURN clob IS
    lcClob clob;
BEGIN
    lcClob := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" mc:Ignorable="w14 w15">
	<w:docDefaults>
		<w:rPrDefault>
			<w:rPr>
				<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>
				<w:sz w:val="22"/>
				<w:szCs w:val="22"/>
				<w:lang w:val="' || grDoc(p_doc_id).lang || '" w:eastAsia="en-US" w:bidi="ar-SA"/>
			</w:rPr>
		</w:rPrDefault>
		<w:pPrDefault>
			<w:pPr>
				<w:spacing w:after="160" w:line="259" w:lineRule="auto"/>
			</w:pPr>
		</w:pPrDefault>
	</w:docDefaults>
	<w:latentStyles w:defLockedState="0" w:defUIPriority="99" w:defSemiHidden="0" w:defUnhideWhenUsed="0" w:defQFormat="0" w:count="371">
		<w:lsdException w:name="Normal" w:uiPriority="0" w:qFormat="1"/>
		<w:lsdException w:name="heading 1" w:uiPriority="9" w:qFormat="1"/>
		<w:lsdException w:name="heading 2" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
		<w:lsdException w:name="heading 3" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
		<w:lsdException w:name="heading 4" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
		<w:lsdException w:name="heading 5" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
		<w:lsdException w:name="heading 6" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
		<w:lsdException w:name="heading 7" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
		<w:lsdException w:name="heading 8" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
		<w:lsdException w:name="heading 9" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
		<w:lsdException w:name="index 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="index 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="index 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="index 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="index 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="index 6" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="index 7" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="index 8" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="index 9" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="toc 1" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="toc 2" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="toc 3" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="toc 4" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="toc 5" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="toc 6" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="toc 7" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="toc 8" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="toc 9" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Normal Indent" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="footnote text" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="annotation text" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="header" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="footer" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="index heading" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="caption" w:semiHidden="1" w:uiPriority="35" w:unhideWhenUsed="1" w:qFormat="1"/>
		<w:lsdException w:name="table of figures" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="envelope address" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="envelope return" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="footnote reference" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="annotation reference" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="line number" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="page number" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="endnote reference" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="endnote text" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="table of authorities" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="macro" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="toa heading" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List Bullet" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List Number" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List Bullet 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List Bullet 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List Bullet 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List Bullet 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List Number 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List Number 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List Number 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List Number 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Title" w:uiPriority="10" w:qFormat="1"/>
		<w:lsdException w:name="Closing" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Signature" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Default Paragraph Font" w:semiHidden="1" w:uiPriority="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Body Text" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Body Text Indent" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List Continue" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List Continue 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List Continue 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List Continue 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List Continue 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Message Header" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Subtitle" w:uiPriority="11" w:qFormat="1"/>
		<w:lsdException w:name="Salutation" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Date" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Body Text First Indent" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Body Text First Indent 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Note Heading" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Body Text 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Body Text 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Body Text Indent 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Body Text Indent 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Block Text" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Hyperlink" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="FollowedHyperlink" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Strong" w:uiPriority="22" w:qFormat="1"/>
		<w:lsdException w:name="Emphasis" w:uiPriority="20" w:qFormat="1"/>
		<w:lsdException w:name="Document Map" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Plain Text" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="E-mail Signature" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="HTML Top of Form" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="HTML Bottom of Form" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Normal (Web)" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="HTML Acronym" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="HTML Address" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="HTML Cite" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="HTML Code" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="HTML Definition" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="HTML Keyboard" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="HTML Preformatted" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="HTML Sample" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="HTML Typewriter" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="HTML Variable" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Normal Table" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="annotation subject" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="No List" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Outline List 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Outline List 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Outline List 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Simple 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Simple 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Simple 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Classic 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Classic 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Classic 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Classic 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Colorful 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Colorful 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Colorful 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Columns 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Columns 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Columns 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Columns 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Columns 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Grid 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Grid 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Grid 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Grid 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Grid 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Grid 6" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Grid 7" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Grid 8" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table List 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table List 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table List 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table List 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table List 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table List 6" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table List 7" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table List 8" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table 3D effects 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table 3D effects 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table 3D effects 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Contemporary" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Elegant" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Professional" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Subtle 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Subtle 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Web 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Web 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Web 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Balloon Text" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Grid" w:uiPriority="39"/>
		<w:lsdException w:name="Table Theme" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Placeholder Text" w:semiHidden="1"/>
		<w:lsdException w:name="No Spacing" w:uiPriority="1" w:qFormat="1"/>
		<w:lsdException w:name="Light Shading" w:uiPriority="60"/>
		<w:lsdException w:name="Light List" w:uiPriority="61"/>
		<w:lsdException w:name="Light Grid" w:uiPriority="62"/>
		<w:lsdException w:name="Medium Shading 1" w:uiPriority="63"/>
		<w:lsdException w:name="Medium Shading 2" w:uiPriority="64"/>
		<w:lsdException w:name="Medium List 1" w:uiPriority="65"/>
		<w:lsdException w:name="Medium List 2" w:uiPriority="66"/>
		<w:lsdException w:name="Medium Grid 1" w:uiPriority="67"/>
		<w:lsdException w:name="Medium Grid 2" w:uiPriority="68"/>
		<w:lsdException w:name="Medium Grid 3" w:uiPriority="69"/>
		<w:lsdException w:name="Dark List" w:uiPriority="70"/>
		<w:lsdException w:name="Colorful Shading" w:uiPriority="71"/>
		<w:lsdException w:name="Colorful List" w:uiPriority="72"/>
		<w:lsdException w:name="Colorful Grid" w:uiPriority="73"/>
		<w:lsdException w:name="Light Shading Accent 1" w:uiPriority="60"/>
		<w:lsdException w:name="Light List Accent 1" w:uiPriority="61"/>
		<w:lsdException w:name="Light Grid Accent 1" w:uiPriority="62"/>
		<w:lsdException w:name="Medium Shading 1 Accent 1" w:uiPriority="63"/>
		<w:lsdException w:name="Medium Shading 2 Accent 1" w:uiPriority="64"/>
		<w:lsdException w:name="Medium List 1 Accent 1" w:uiPriority="65"/>
		<w:lsdException w:name="Revision" w:semiHidden="1"/>
		<w:lsdException w:name="List Paragraph" w:uiPriority="34" w:qFormat="1"/>
		<w:lsdException w:name="Quote" w:uiPriority="29" w:qFormat="1"/>
		<w:lsdException w:name="Intense Quote" w:uiPriority="30" w:qFormat="1"/>
		<w:lsdException w:name="Medium List 2 Accent 1" w:uiPriority="66"/>
		<w:lsdException w:name="Medium Grid 1 Accent 1" w:uiPriority="67"/>
		<w:lsdException w:name="Medium Grid 2 Accent 1" w:uiPriority="68"/>
		<w:lsdException w:name="Medium Grid 3 Accent 1" w:uiPriority="69"/>
		<w:lsdException w:name="Dark List Accent 1" w:uiPriority="70"/>
		<w:lsdException w:name="Colorful Shading Accent 1" w:uiPriority="71"/>
		<w:lsdException w:name="Colorful List Accent 1" w:uiPriority="72"/>
		<w:lsdException w:name="Colorful Grid Accent 1" w:uiPriority="73"/>
		<w:lsdException w:name="Light Shading Accent 2" w:uiPriority="60"/>
		<w:lsdException w:name="Light List Accent 2" w:uiPriority="61"/>
		<w:lsdException w:name="Light Grid Accent 2" w:uiPriority="62"/>
		<w:lsdException w:name="Medium Shading 1 Accent 2" w:uiPriority="63"/>
		<w:lsdException w:name="Medium Shading 2 Accent 2" w:uiPriority="64"/>
		<w:lsdException w:name="Medium List 1 Accent 2" w:uiPriority="65"/>
		<w:lsdException w:name="Medium List 2 Accent 2" w:uiPriority="66"/>
		<w:lsdException w:name="Medium Grid 1 Accent 2" w:uiPriority="67"/>
		<w:lsdException w:name="Medium Grid 2 Accent 2" w:uiPriority="68"/>
		<w:lsdException w:name="Medium Grid 3 Accent 2" w:uiPriority="69"/>
		<w:lsdException w:name="Dark List Accent 2" w:uiPriority="70"/>
		<w:lsdException w:name="Colorful Shading Accent 2" w:uiPriority="71"/>
		<w:lsdException w:name="Colorful List Accent 2" w:uiPriority="72"/>
		<w:lsdException w:name="Colorful Grid Accent 2" w:uiPriority="73"/>
		<w:lsdException w:name="Light Shading Accent 3" w:uiPriority="60"/>
		<w:lsdException w:name="Light List Accent 3" w:uiPriority="61"/>
		<w:lsdException w:name="Light Grid Accent 3" w:uiPriority="62"/>
		<w:lsdException w:name="Medium Shading 1 Accent 3" w:uiPriority="63"/>
		<w:lsdException w:name="Medium Shading 2 Accent 3" w:uiPriority="64"/>
		<w:lsdException w:name="Medium List 1 Accent 3" w:uiPriority="65"/>
		<w:lsdException w:name="Medium List 2 Accent 3" w:uiPriority="66"/>
		<w:lsdException w:name="Medium Grid 1 Accent 3" w:uiPriority="67"/>
		<w:lsdException w:name="Medium Grid 2 Accent 3" w:uiPriority="68"/>
		<w:lsdException w:name="Medium Grid 3 Accent 3" w:uiPriority="69"/>
		<w:lsdException w:name="Dark List Accent 3" w:uiPriority="70"/>
		<w:lsdException w:name="Colorful Shading Accent 3" w:uiPriority="71"/>
		<w:lsdException w:name="Colorful List Accent 3" w:uiPriority="72"/>
		<w:lsdException w:name="Colorful Grid Accent 3" w:uiPriority="73"/>
		<w:lsdException w:name="Light Shading Accent 4" w:uiPriority="60"/>
		<w:lsdException w:name="Light List Accent 4" w:uiPriority="61"/>
		<w:lsdException w:name="Light Grid Accent 4" w:uiPriority="62"/>
		<w:lsdException w:name="Medium Shading 1 Accent 4" w:uiPriority="63"/>
		<w:lsdException w:name="Medium Shading 2 Accent 4" w:uiPriority="64"/>
		<w:lsdException w:name="Medium List 1 Accent 4" w:uiPriority="65"/>
		<w:lsdException w:name="Medium List 2 Accent 4" w:uiPriority="66"/>
		<w:lsdException w:name="Medium Grid 1 Accent 4" w:uiPriority="67"/>
		<w:lsdException w:name="Medium Grid 2 Accent 4" w:uiPriority="68"/>
		<w:lsdException w:name="Medium Grid 3 Accent 4" w:uiPriority="69"/>
		<w:lsdException w:name="Dark List Accent 4" w:uiPriority="70"/>
		<w:lsdException w:name="Colorful Shading Accent 4" w:uiPriority="71"/>
		<w:lsdException w:name="Colorful List Accent 4" w:uiPriority="72"/>
		<w:lsdException w:name="Colorful Grid Accent 4" w:uiPriority="73"/>
		<w:lsdException w:name="Light Shading Accent 5" w:uiPriority="60"/>
		<w:lsdException w:name="Light List Accent 5" w:uiPriority="61"/>
		<w:lsdException w:name="Light Grid Accent 5" w:uiPriority="62"/>
		<w:lsdException w:name="Medium Shading 1 Accent 5" w:uiPriority="63"/>
		<w:lsdException w:name="Medium Shading 2 Accent 5" w:uiPriority="64"/>
		<w:lsdException w:name="Medium List 1 Accent 5" w:uiPriority="65"/>
		<w:lsdException w:name="Medium List 2 Accent 5" w:uiPriority="66"/>
		<w:lsdException w:name="Medium Grid 1 Accent 5" w:uiPriority="67"/>
		<w:lsdException w:name="Medium Grid 2 Accent 5" w:uiPriority="68"/>
		<w:lsdException w:name="Medium Grid 3 Accent 5" w:uiPriority="69"/>
		<w:lsdException w:name="Dark List Accent 5" w:uiPriority="70"/>
		<w:lsdException w:name="Colorful Shading Accent 5" w:uiPriority="71"/>
		<w:lsdException w:name="Colorful List Accent 5" w:uiPriority="72"/>
		<w:lsdException w:name="Colorful Grid Accent 5" w:uiPriority="73"/>
		<w:lsdException w:name="Light Shading Accent 6" w:uiPriority="60"/>
		<w:lsdException w:name="Light List Accent 6" w:uiPriority="61"/>
		<w:lsdException w:name="Light Grid Accent 6" w:uiPriority="62"/>
		<w:lsdException w:name="Medium Shading 1 Accent 6" w:uiPriority="63"/>
		<w:lsdException w:name="Medium Shading 2 Accent 6" w:uiPriority="64"/>
		<w:lsdException w:name="Medium List 1 Accent 6" w:uiPriority="65"/>
		<w:lsdException w:name="Medium List 2 Accent 6" w:uiPriority="66"/>
		<w:lsdException w:name="Medium Grid 1 Accent 6" w:uiPriority="67"/>
		<w:lsdException w:name="Medium Grid 2 Accent 6" w:uiPriority="68"/>
		<w:lsdException w:name="Medium Grid 3 Accent 6" w:uiPriority="69"/>
		<w:lsdException w:name="Dark List Accent 6" w:uiPriority="70"/>
		<w:lsdException w:name="Colorful Shading Accent 6" w:uiPriority="71"/>
		<w:lsdException w:name="Colorful List Accent 6" w:uiPriority="72"/>
		<w:lsdException w:name="Colorful Grid Accent 6" w:uiPriority="73"/>
		<w:lsdException w:name="Subtle Emphasis" w:uiPriority="19" w:qFormat="1"/>
		<w:lsdException w:name="Intense Emphasis" w:uiPriority="21" w:qFormat="1"/>
		<w:lsdException w:name="Subtle Reference" w:uiPriority="31" w:qFormat="1"/>
		<w:lsdException w:name="Intense Reference" w:uiPriority="32" w:qFormat="1"/>
		<w:lsdException w:name="Book Title" w:uiPriority="33" w:qFormat="1"/>
		<w:lsdException w:name="Bibliography" w:semiHidden="1" w:uiPriority="37" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="TOC Heading" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1" w:qFormat="1"/>
		<w:lsdException w:name="Plain Table 1" w:uiPriority="41"/>
		<w:lsdException w:name="Plain Table 2" w:uiPriority="42"/>
		<w:lsdException w:name="Plain Table 3" w:uiPriority="43"/>
		<w:lsdException w:name="Plain Table 4" w:uiPriority="44"/>
		<w:lsdException w:name="Plain Table 5" w:uiPriority="45"/>
		<w:lsdException w:name="Grid Table Light" w:uiPriority="40"/>
		<w:lsdException w:name="Grid Table 1 Light" w:uiPriority="46"/>
		<w:lsdException w:name="Grid Table 2" w:uiPriority="47"/>
		<w:lsdException w:name="Grid Table 3" w:uiPriority="48"/>
		<w:lsdException w:name="Grid Table 4" w:uiPriority="49"/>
		<w:lsdException w:name="Grid Table 5 Dark" w:uiPriority="50"/>
		<w:lsdException w:name="Grid Table 6 Colorful" w:uiPriority="51"/>
		<w:lsdException w:name="Grid Table 7 Colorful" w:uiPriority="52"/>
		<w:lsdException w:name="Grid Table 1 Light Accent 1" w:uiPriority="46"/>
		<w:lsdException w:name="Grid Table 2 Accent 1" w:uiPriority="47"/>
		<w:lsdException w:name="Grid Table 3 Accent 1" w:uiPriority="48"/>
		<w:lsdException w:name="Grid Table 4 Accent 1" w:uiPriority="49"/>
		<w:lsdException w:name="Grid Table 5 Dark Accent 1" w:uiPriority="50"/>
		<w:lsdException w:name="Grid Table 6 Colorful Accent 1" w:uiPriority="51"/>
		<w:lsdException w:name="Grid Table 7 Colorful Accent 1" w:uiPriority="52"/>
		<w:lsdException w:name="Grid Table 1 Light Accent 2" w:uiPriority="46"/>
		<w:lsdException w:name="Grid Table 2 Accent 2" w:uiPriority="47"/>
		<w:lsdException w:name="Grid Table 3 Accent 2" w:uiPriority="48"/>
		<w:lsdException w:name="Grid Table 4 Accent 2" w:uiPriority="49"/>
		<w:lsdException w:name="Grid Table 5 Dark Accent 2" w:uiPriority="50"/>
		<w:lsdException w:name="Grid Table 6 Colorful Accent 2" w:uiPriority="51"/>
		<w:lsdException w:name="Grid Table 7 Colorful Accent 2" w:uiPriority="52"/>
		<w:lsdException w:name="Grid Table 1 Light Accent 3" w:uiPriority="46"/>
		<w:lsdException w:name="Grid Table 2 Accent 3" w:uiPriority="47"/>
		<w:lsdException w:name="Grid Table 3 Accent 3" w:uiPriority="48"/>
		<w:lsdException w:name="Grid Table 4 Accent 3" w:uiPriority="49"/>
		<w:lsdException w:name="Grid Table 5 Dark Accent 3" w:uiPriority="50"/>
		<w:lsdException w:name="Grid Table 6 Colorful Accent 3" w:uiPriority="51"/>
		<w:lsdException w:name="Grid Table 7 Colorful Accent 3" w:uiPriority="52"/>
		<w:lsdException w:name="Grid Table 1 Light Accent 4" w:uiPriority="46"/>
		<w:lsdException w:name="Grid Table 2 Accent 4" w:uiPriority="47"/>
		<w:lsdException w:name="Grid Table 3 Accent 4" w:uiPriority="48"/>
		<w:lsdException w:name="Grid Table 4 Accent 4" w:uiPriority="49"/>
		<w:lsdException w:name="Grid Table 5 Dark Accent 4" w:uiPriority="50"/>
		<w:lsdException w:name="Grid Table 6 Colorful Accent 4" w:uiPriority="51"/>
		<w:lsdException w:name="Grid Table 7 Colorful Accent 4" w:uiPriority="52"/>
		<w:lsdException w:name="Grid Table 1 Light Accent 5" w:uiPriority="46"/>
		<w:lsdException w:name="Grid Table 2 Accent 5" w:uiPriority="47"/>
		<w:lsdException w:name="Grid Table 3 Accent 5" w:uiPriority="48"/>
		<w:lsdException w:name="Grid Table 4 Accent 5" w:uiPriority="49"/>
		<w:lsdException w:name="Grid Table 5 Dark Accent 5" w:uiPriority="50"/>
		<w:lsdException w:name="Grid Table 6 Colorful Accent 5" w:uiPriority="51"/>
		<w:lsdException w:name="Grid Table 7 Colorful Accent 5" w:uiPriority="52"/>
		<w:lsdException w:name="Grid Table 1 Light Accent 6" w:uiPriority="46"/>
		<w:lsdException w:name="Grid Table 2 Accent 6" w:uiPriority="47"/>
		<w:lsdException w:name="Grid Table 3 Accent 6" w:uiPriority="48"/>
		<w:lsdException w:name="Grid Table 4 Accent 6" w:uiPriority="49"/>
		<w:lsdException w:name="Grid Table 5 Dark Accent 6" w:uiPriority="50"/>
		<w:lsdException w:name="Grid Table 6 Colorful Accent 6" w:uiPriority="51"/>
		<w:lsdException w:name="Grid Table 7 Colorful Accent 6" w:uiPriority="52"/>
		<w:lsdException w:name="List Table 1 Light" w:uiPriority="46"/>
		<w:lsdException w:name="List Table 2" w:uiPriority="47"/>
		<w:lsdException w:name="List Table 3" w:uiPriority="48"/>
		<w:lsdException w:name="List Table 4" w:uiPriority="49"/>
		<w:lsdException w:name="List Table 5 Dark" w:uiPriority="50"/>
		<w:lsdException w:name="List Table 6 Colorful" w:uiPriority="51"/>
		<w:lsdException w:name="List Table 7 Colorful" w:uiPriority="52"/>
		<w:lsdException w:name="List Table 1 Light Accent 1" w:uiPriority="46"/>
		<w:lsdException w:name="List Table 2 Accent 1" w:uiPriority="47"/>
		<w:lsdException w:name="List Table 3 Accent 1" w:uiPriority="48"/>
		<w:lsdException w:name="List Table 4 Accent 1" w:uiPriority="49"/>
		<w:lsdException w:name="List Table 5 Dark Accent 1" w:uiPriority="50"/>
		<w:lsdException w:name="List Table 6 Colorful Accent 1" w:uiPriority="51"/>
		<w:lsdException w:name="List Table 7 Colorful Accent 1" w:uiPriority="52"/>
		<w:lsdException w:name="List Table 1 Light Accent 2" w:uiPriority="46"/>
		<w:lsdException w:name="List Table 2 Accent 2" w:uiPriority="47"/>
		<w:lsdException w:name="List Table 3 Accent 2" w:uiPriority="48"/>
		<w:lsdException w:name="List Table 4 Accent 2" w:uiPriority="49"/>
		<w:lsdException w:name="List Table 5 Dark Accent 2" w:uiPriority="50"/>
		<w:lsdException w:name="List Table 6 Colorful Accent 2" w:uiPriority="51"/>
		<w:lsdException w:name="List Table 7 Colorful Accent 2" w:uiPriority="52"/>
		<w:lsdException w:name="List Table 1 Light Accent 3" w:uiPriority="46"/>
		<w:lsdException w:name="List Table 2 Accent 3" w:uiPriority="47"/>
		<w:lsdException w:name="List Table 3 Accent 3" w:uiPriority="48"/>
		<w:lsdException w:name="List Table 4 Accent 3" w:uiPriority="49"/>
		<w:lsdException w:name="List Table 5 Dark Accent 3" w:uiPriority="50"/>
		<w:lsdException w:name="List Table 6 Colorful Accent 3" w:uiPriority="51"/>
		<w:lsdException w:name="List Table 7 Colorful Accent 3" w:uiPriority="52"/>
		<w:lsdException w:name="List Table 1 Light Accent 4" w:uiPriority="46"/>
		<w:lsdException w:name="List Table 2 Accent 4" w:uiPriority="47"/>
		<w:lsdException w:name="List Table 3 Accent 4" w:uiPriority="48"/>
		<w:lsdException w:name="List Table 4 Accent 4" w:uiPriority="49"/>
		<w:lsdException w:name="List Table 5 Dark Accent 4" w:uiPriority="50"/>
		<w:lsdException w:name="List Table 6 Colorful Accent 4" w:uiPriority="51"/>
		<w:lsdException w:name="List Table 7 Colorful Accent 4" w:uiPriority="52"/>
		<w:lsdException w:name="List Table 1 Light Accent 5" w:uiPriority="46"/>
		<w:lsdException w:name="List Table 2 Accent 5" w:uiPriority="47"/>
		<w:lsdException w:name="List Table 3 Accent 5" w:uiPriority="48"/>
		<w:lsdException w:name="List Table 4 Accent 5" w:uiPriority="49"/>
		<w:lsdException w:name="List Table 5 Dark Accent 5" w:uiPriority="50"/>
		<w:lsdException w:name="List Table 6 Colorful Accent 5" w:uiPriority="51"/>
		<w:lsdException w:name="List Table 7 Colorful Accent 5" w:uiPriority="52"/>
		<w:lsdException w:name="List Table 1 Light Accent 6" w:uiPriority="46"/>
		<w:lsdException w:name="List Table 2 Accent 6" w:uiPriority="47"/>
		<w:lsdException w:name="List Table 3 Accent 6" w:uiPriority="48"/>
		<w:lsdException w:name="List Table 4 Accent 6" w:uiPriority="49"/>
		<w:lsdException w:name="List Table 5 Dark Accent 6" w:uiPriority="50"/>
		<w:lsdException w:name="List Table 6 Colorful Accent 6" w:uiPriority="51"/>
		<w:lsdException w:name="List Table 7 Colorful Accent 6" w:uiPriority="52"/>';
		
	lcClob := lcClob || '
	</w:latentStyles>
	<w:style w:type="paragraph" w:default="1" w:styleId="Normal">
		<w:name w:val="Normal"/>
		<w:qFormat/>
	</w:style>
	<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">
		<w:name w:val="Default Paragraph Font"/>
		<w:uiPriority w:val="1"/>
		<w:semiHidden/>
		<w:unhideWhenUsed/>
	</w:style>
	<w:style w:type="table" w:default="1" w:styleId="TableNormal">
		<w:name w:val="Normal Table"/>
		<w:uiPriority w:val="99"/>
		<w:semiHidden/>
		<w:unhideWhenUsed/>
		<w:tblPr>
			<w:tblInd w:w="0" w:type="dxa"/>
			<w:tblCellMar>
				<w:top w:w="0" w:type="dxa"/>
				<w:left w:w="108" w:type="dxa"/>
				<w:bottom w:w="0" w:type="dxa"/>
				<w:right w:w="108" w:type="dxa"/>
			</w:tblCellMar>
		</w:tblPr>
	</w:style>
	<w:style w:type="numbering" w:default="1" w:styleId="NoList">
		<w:name w:val="No List"/>
		<w:uiPriority w:val="99"/>
		<w:semiHidden/>
		<w:unhideWhenUsed/>
	</w:style>
	<w:style w:type="table" w:styleId="TableGrid">
		<w:name w:val="Table Grid"/>
		<w:basedOn w:val="TableNormal"/>
		<w:uiPriority w:val="39"/>
		<w:rsid w:val="00B34BE1"/>
		<w:pPr>
			<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
		</w:pPr>
		<w:tblPr>
			<w:tblBorders>
				<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
				<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
				<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
				<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
				<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
				<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
			</w:tblBorders>
		</w:tblPr>
	</w:style>
	<w:style w:type="paragraph" w:styleId="Heading1">
		<w:name w:val="heading 1"/>
		<w:basedOn w:val="Normal"/>
		<w:next w:val="Normal"/>
		<w:link w:val="Heading1Char"/>
		<w:uiPriority w:val="9"/>
		<w:qFormat/>
		<w:pPr>
			<w:keepNext/>
			<w:keepLines/>
			<w:spacing w:before="240" w:after="0"/>
			<w:outlineLvl w:val="0"/>
		</w:pPr>
		<w:rPr>
			<w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
			<w:color w:val="2E74B5" w:themeColor="accent1" w:themeShade="BF"/>
			<w:sz w:val="32"/>
			<w:szCs w:val="32"/>
		</w:rPr>
	</w:style>
	<w:style w:type="paragraph" w:styleId="Heading2">
		<w:name w:val="heading 2"/>
		<w:basedOn w:val="Normal"/>
		<w:next w:val="Normal"/>
		<w:link w:val="Heading2Char"/>
		<w:uiPriority w:val="9"/>
		<w:unhideWhenUsed/>
		<w:qFormat/>
		<w:rsid w:val="00E06CAF"/>
		<w:pPr>
			<w:keepNext/>
			<w:keepLines/>
			<w:spacing w:before="40" w:after="0"/>
			<w:outlineLvl w:val="1"/>
		</w:pPr>
		<w:rPr>
			<w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
			<w:color w:val="2E74B5" w:themeColor="accent1" w:themeShade="BF"/>
			<w:sz w:val="26"/>
			<w:szCs w:val="26"/>
		</w:rPr>
	</w:style>
	<w:style w:type="paragraph" w:styleId="Heading3">
		<w:name w:val="heading 3"/>
		<w:basedOn w:val="Normal"/>
		<w:next w:val="Normal"/>
		<w:link w:val="Heading3Char"/>
		<w:uiPriority w:val="9"/>
		<w:unhideWhenUsed/>
		<w:qFormat/>
		<w:rsid w:val="00FA6340"/>
		<w:pPr>
			<w:keepNext/>
			<w:keepLines/>
			<w:spacing w:before="40" w:after="0"/>
			<w:outlineLvl w:val="2"/>
		</w:pPr>
		<w:rPr>
			<w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
			<w:color w:val="1F4D78" w:themeColor="accent1" w:themeShade="7F"/>
			<w:sz w:val="24"/>
			<w:szCs w:val="24"/>
		</w:rPr>
	</w:style>
	<w:style w:type="paragraph" w:styleId="Heading4">
		<w:name w:val="heading 4"/>
		<w:basedOn w:val="Normal"/>
		<w:next w:val="Normal"/>
		<w:link w:val="Heading4Char"/>
		<w:uiPriority w:val="9"/>
		<w:unhideWhenUsed/>
		<w:qFormat/>
		<w:rsid w:val="00FA6340"/>
		<w:pPr>
			<w:keepNext/>
			<w:keepLines/>
			<w:spacing w:before="40" w:after="0"/>
			<w:outlineLvl w:val="3"/>
		</w:pPr>
		<w:rPr>
			<w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
			<w:i/>
			<w:iCs/>
			<w:color w:val="2E74B5" w:themeColor="accent1" w:themeShade="BF"/>
		</w:rPr>
	</w:style>
	<w:style w:type="paragraph" w:styleId="ListParagraph">
		<w:name w:val="List Paragraph"/>
		<w:basedOn w:val="Normal"/>
		<w:uiPriority w:val="34"/>
		<w:qFormat/>
		<w:rsid w:val="00476CD7"/>
		<w:pPr>
			<w:ind w:left="720"/>
			<w:contextualSpacing/>
		</w:pPr>
	</w:style>
</w:styles>';

    RETURN lcClob;
END f_styles;


FUNCTION f_web_settings RETURN clob IS
BEGIN
    RETURN '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:webSettings xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" mc:Ignorable="w14 w15">
	<w:optimizeForBrowser/>
	<w:allowPNG/>
</w:webSettings>';
END f_web_settings;


FUNCTION f_numbering(p_doc_id pls_integer) RETURN clob IS
    lcClob clob;
BEGIN
    if grDoc(p_doc_id).lists.count = 0 then
        RETURN null;
    end if;
    
    lcClob := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 wp14">
';

    FOR t IN grDoc(p_doc_id).lists.first .. grDoc(p_doc_id).lists.last LOOP
        lcClob := lcClob || '	<w:abstractNum w:abstractNumId="' || (t - 1) || '" w15:restartNumberingAfterBreak="0">
		<w:multiLevelType w:val="hybridMultilevel"/>
		<w:lvl w:ilvl="0">
			<w:start w:val="' || nvl(grDoc(p_doc_id).lists(t).num_start_value, '1') || '"/>
			<w:numFmt w:val="' || grDoc(p_doc_id).lists(t).list_type || '"/>
			<w:lvlText w:val="' || nvl(grDoc(p_doc_id).lists(t).bullet_char, '%1.')  || '"/>
			<w:lvlJc w:val="left"/>
			<w:pPr>
				<w:ind w:left="720" w:hanging="360"/>
			</w:pPr>
			<w:rPr>
				<w:rFonts ' || (CASE WHEN grDoc(p_doc_id).lists(t).bullet_font is null THEN null ELSE 'w:ascii="' || grDoc(p_doc_id).lists(t).bullet_font || '" w:hAnsi="' || grDoc(p_doc_id).lists(t).bullet_font || '" ' END)  || 'w:hint="default"/>
			</w:rPr>
		</w:lvl>
		<w:lvl w:ilvl="1" w:tentative="1">
			<w:start w:val="1"/>
			<w:numFmt w:val="lowerLetter"/>
			<w:lvlText w:val="%2."/>
			<w:lvlJc w:val="left"/>
			<w:pPr>
				<w:ind w:left="1440" w:hanging="360"/>
			</w:pPr>
		</w:lvl>
		<w:lvl w:ilvl="2" w:tentative="1">
			<w:start w:val="1"/>
			<w:numFmt w:val="lowerRoman"/>
			<w:lvlText w:val="%3."/>
			<w:lvlJc w:val="right"/>
			<w:pPr>
				<w:ind w:left="2160" w:hanging="180"/>
			</w:pPr>
		</w:lvl>
		<w:lvl w:ilvl="3" w:tentative="1">
			<w:start w:val="1"/>
			<w:numFmt w:val="decimal"/>
			<w:lvlText w:val="%4."/>
			<w:lvlJc w:val="left"/>
			<w:pPr>
				<w:ind w:left="2880" w:hanging="360"/>
			</w:pPr>
		</w:lvl>
		<w:lvl w:ilvl="4" w:tentative="1">
			<w:start w:val="1"/>
			<w:numFmt w:val="lowerLetter"/>
			<w:lvlText w:val="%5."/>
			<w:lvlJc w:val="left"/>
			<w:pPr>
				<w:ind w:left="3600" w:hanging="360"/>
			</w:pPr>
		</w:lvl>
		<w:lvl w:ilvl="5" w:tentative="1">
			<w:start w:val="1"/>
			<w:numFmt w:val="lowerRoman"/>
			<w:lvlText w:val="%6."/>
			<w:lvlJc w:val="right"/>
			<w:pPr>
				<w:ind w:left="4320" w:hanging="180"/>
			</w:pPr>
		</w:lvl>
		<w:lvl w:ilvl="6" w:tentative="1">
			<w:start w:val="1"/>
			<w:numFmt w:val="decimal"/>
			<w:lvlText w:val="%7."/>
			<w:lvlJc w:val="left"/>
			<w:pPr>
				<w:ind w:left="5040" w:hanging="360"/>
			</w:pPr>
		</w:lvl>
		<w:lvl w:ilvl="7" w:tentative="1">
			<w:start w:val="1"/>
			<w:numFmt w:val="lowerLetter"/>
			<w:lvlText w:val="%8."/>
			<w:lvlJc w:val="left"/>
			<w:pPr>
				<w:ind w:left="5760" w:hanging="360"/>
			</w:pPr>
		</w:lvl>
		<w:lvl w:ilvl="8" w:tentative="1">
			<w:start w:val="1"/>
			<w:numFmt w:val="lowerRoman"/>
			<w:lvlText w:val="%9."/>
			<w:lvlJc w:val="right"/>
			<w:pPr>
				<w:ind w:left="6480" w:hanging="180"/>
			</w:pPr>
		</w:lvl>
	</w:abstractNum>
';
    END LOOP;
    
    FOR t IN grDoc(p_doc_id).lists.first .. grDoc(p_doc_id).lists.last LOOP
        lcClob := lcClob || '	<w:num w:numId="' || t || '"><w:abstractNumId w:val="' || (t - 1) || '"/></w:num>
';
    END LOOP;

    lcClob := lcClob || '</w:numbering>';
    
    RETURN lcClob;
    
END f_numbering;




PROCEDURE p_add_clob_text(p_text clob) IS
BEGIN
    gcClob := gcClob || p_text || chr(10);
END p_add_clob_text;


PROCEDURE p_xml_image(
    p_doc_id number,
    p_image_data r_image_data,
    p_paragraph_yn varchar2
    ) IS
    
    lcImageName varchar2(100) := grDoc(p_doc_id).images(p_image_data.image_id).image_name;
    lcImageDesc varchar2(1000) := 'image from document';
    lnWidth pls_integer := f_unit_convert(p_doc_id, 'image', p_image_data.width);
    lnHeight pls_integer := f_unit_convert(p_doc_id, 'image', p_image_data.height);
    lnRotateAngle pls_integer := f_unit_convert(p_doc_id, 'image_rotate', p_image_data.rotate_angle);
    lcInlineAnchor varchar2(32000) := '<wp:inline distT="0" distB="0" distL="0" distR="0">';
    lcPositionH varchar2(20) := to_char(f_unit_convert(p_doc_id, 'image', p_image_data.position_h));
    lcPositionV varchar2(20) := to_char(f_unit_convert(p_doc_id, 'image', p_image_data.position_v));
    
BEGIN
    if p_paragraph_yn = 'Y' then
        p_add_clob_text('<w:p>');
    end if;

    --for anchored images
    if p_image_data.inline_yn = 'N' then
        lcInlineAnchor := '
        <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251656704" behindDoc="0" locked="1" layoutInCell="1" allowOverlap="0">
            <wp:simplePos x="360000" y="0"/>
            <wp:positionH relativeFrom="' || p_image_data.relative_from_h || '">
                <wp:' || p_image_data.position_type_h || '>' || (CASE p_image_data.position_type_h WHEN 'align' THEN p_image_data.position_align_h ELSE lcPositionH END) || '</wp:' || p_image_data.position_type_h || '>
            </wp:positionH>
            <wp:positionV relativeFrom="' || p_image_data.relative_from_v || '">
                <wp:' || p_image_data.position_type_v || '>' || (CASE p_image_data.position_type_v WHEN 'align' THEN p_image_data.position_align_v ELSE lcPositionV END) || '</wp:' || p_image_data.position_type_v || '>
            </wp:positionV>
        ';
    end if;

    p_add_clob_text('
			<w:r>
				<w:drawing>
					' || lcInlineAnchor || '
						<wp:extent cx="' || lnWidth || '" cy="' || lnHeight || '"/>
						<wp:effectExtent 
                            l="' || f_unit_convert(p_doc_id, 'image', p_image_data.extent_area_left) || '" 
                            t="' || f_unit_convert(p_doc_id, 'image', p_image_data.extent_area_top) || '" 
                            r="' || f_unit_convert(p_doc_id, 'image', p_image_data.extent_area_right) || '" 
                            b="' || f_unit_convert(p_doc_id, 'image', p_image_data.extent_area_bottom) || '"/>
						<wp:wrapNone/>
						<wp:docPr id="' || p_image_data.image_id || '" name="' || lcImageName || '" descr="' || lcImageDesc || '"/>
						<wp:cNvGraphicFramePr>
							<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
						</wp:cNvGraphicFramePr>
						<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
							<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
								<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
									<pic:nvPicPr>
										<pic:cNvPr id="' || p_image_data.image_id || '" name="' || lcImageName || '" descr="' || lcImageDesc || '"/>
										<pic:cNvPicPr>
											<a:picLocks noChangeAspect="1" noChangeArrowheads="1"/>
										</pic:cNvPicPr>
									</pic:nvPicPr>
									<pic:blipFill>
										<a:blip r:embed="rId' || grDoc(p_doc_id).images(p_image_data.image_id).rel_id || '" cstate="print">
											<a:extLst>
												<a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
													<a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>
												</a:ext>
											</a:extLst>
										</a:blip>
										<a:srcRect/>
										<a:stretch>
											<a:fillRect/>
										</a:stretch>
									</pic:blipFill>
									<pic:spPr bwMode="auto">
										<a:xfrm rot="' || lnRotateAngle || '">
											<a:off x="0" y="0"/>
											<a:ext cx="' || lnWidth || '" cy="' || lnHeight || '"/>
										</a:xfrm>
										<a:prstGeom prst="rect">
											<a:avLst/>
										</a:prstGeom>
										<a:noFill/>
										<a:ln>
											<a:noFill/>
										</a:ln>
									</pic:spPr>
								</pic:pic>
							</a:graphicData>
						</a:graphic>
					</wp:' || (CASE p_image_data.inline_yn WHEN 'Y' THEN 'inline' ELSE 'anchor' END) || '>
				</w:drawing>
			</w:r>
');

    if p_paragraph_yn = 'Y' then
        p_add_clob_text('</w:p>');
    end if;

END p_xml_image;


PROCEDURE p_xml_text(
    p_doc_id pls_integer, 
    p_text r_text
    ) IS

    CURSOR c_parts IS
    SELECT
        REGEXP_SUBSTR(p_text.text, '[^' || p_text.newline_character || ']+', 1, LEVEL) as text_part 
    FROM DUAL
    CONNECT BY 
        REGEXP_SUBSTR(p_text.text, '[^' || p_text.newline_character || ']+', 1, LEVEL) IS NOT NULL
    ;
    
    TYPE t_parts IS TABLE OF c_parts%ROWTYPE;
    l_parts t_parts;
    

BEGIN
    if p_text.text is not null then
        p_add_clob_text('<w:r>');

        if not p_text.font.from_paragraph then  
            p_add_clob_text('<w:rPr>');
            
            p_add_clob_text('<w:rFonts w:ascii="' || p_text.font.font_name || '" w:hAnsi="' || p_text.font.font_name || '"/>');
            p_add_clob_text('<w:sz w:val="' || (p_text.font.font_size * 2) || '"/>');
            p_add_clob_text('<w:szCs w:val="' || (p_text.font.font_size * 2) || '"/>');
            p_add_clob_text('<w:color w:val="' || p_text.font.color || '"/>');

            if p_text.font.bold then
                p_add_clob_text('<w:b/>');
            end if;

            if p_text.font.italic then
                p_add_clob_text('<w:i/>');
            end if;

            if p_text.font.underline then
                p_add_clob_text('<w:u w:val="single"/>');
            end if;

            p_add_clob_text('</w:rPr>');
            
        end if;

        if p_text.replace_newline then  --replace newline character with Word's actual newline tag

            OPEN c_parts;
            FETCH c_parts BULK COLLECT INTO l_parts;
            CLOSE c_parts;
            
            FOR t IN 1 .. l_parts.count LOOP
                
                p_add_clob_text('<w:t xml:space="preserve">' || dbms_xmlgen.convert(l_parts(t).text_part) || '</w:t>');
                
                if t < l_parts.count then
                    p_add_clob_text('<w:br/>');
                end if;
            
            END LOOP;
        
        else
            p_add_clob_text('<w:t xml:space="preserve">' || dbms_xmlgen.convert(p_text.text) || '</w:t>');
            
        end if;

        p_add_clob_text('</w:r>');

    elsif p_text.line_break = true then
        p_add_clob_text('<w:br/>');
        
    elsif p_text.image_data.image_id is not null then
        p_xml_image(
            p_doc_id => p_doc_id,
            p_image_data => p_text.image_data,
            p_paragraph_yn => 'N'
        );
    end if;
    
END p_xml_text;


PROCEDURE p_xml_paragraph(
    p_doc_id pls_integer, 
    p_paragraph r_paragraph
    ) IS
BEGIN
    --start
    p_add_clob_text('<w:p>');
    p_add_clob_text('<w:pPr>');

    --alignment
    if p_paragraph.alignment_h <> 'LEFT' then
        p_add_clob_text('<w:jc w:val="' || lower(p_paragraph.alignment_h) || '"/>');
    end if;

    --style (if set)
    if p_paragraph.style is not null and p_paragraph.list_id is null then
        p_add_clob_text('<w:pStyle w:val="' || p_paragraph.style || '"/>');
    end if;

    --numbering
    if p_paragraph.list_id is not null then
        p_add_clob_text('<w:pStyle w:val="ListParagraph"/>
<w:numPr>
<w:ilvl w:val="0"/>
<w:numId w:val="' || p_paragraph.list_id || '"/>
</w:numPr>');
    end if;

    p_add_clob_text('</w:pPr>');
    
    --texts
    if p_paragraph.texts is not null then
        FOR t IN p_paragraph.texts.first .. p_paragraph.texts.last LOOP
            p_xml_text(p_doc_id, p_paragraph.texts(t));
        END LOOP;
    end if;
    
    --end
    p_add_clob_text('</w:p>');
END p_xml_paragraph;

PROCEDURE p_xml_table(
    p_doc_id pls_integer, 
    p_table r_table
    ) IS
    
    lnWidth pls_integer := 0;
    lcClobTop clob;
    lcClobBottom clob;
    lcClobLeft clob;
    lcClobRight clob;
    lcClobInsideH clob;
    lcClobInsideV clob;

    FUNCTION f_border_text(
        p_which varchar2, 
        p_border r_border) RETURN clob IS
        lcClob clob;
    BEGIN
        if p_border.border_type is not null then
            lcClob := '<w:' || p_which || ' ';
            
            lcClob := lcClob || 'w:val="' || p_border.border_type || '" ';

            lcClob := lcClob || 'w:sz="' || ((CASE WHEN lower(p_border.border_type) = 'none' THEN 0 ELSE nvl(p_border.width, 0.5) END) * 8) || '" ';

            lcClob := lcClob || 'w:space="0" w:color="' || nvl(p_border.color, 'auto') || '"/>';
        end if;
        
        RETURN lcClob;
    END;
    
BEGIN
    --table start
    p_add_clob_text('<w:tbl>
    <w:tblPr>
        <w:tblStyle w:val="TableGrid"/>');
        
    --table width
    if p_table.width is null then
        p_add_clob_text('<w:tblW w:w="0" w:type="auto"/>');
    else
        p_add_clob_text('<w:tblW w:w="' || p_table.width || '" w:type="dxa"/>');
    end if;
        
    --table borders
    lcClobTop := f_border_text('top', p_table.border_top);
    lcClobBottom := f_border_text('bottom', p_table.border_bottom);
    lcClobLeft := f_border_text('left', p_table.border_left);
    lcClobRight := f_border_text('right', p_table.border_right);
    lcClobInsideH := f_border_text('insideH', p_table.border_inside_h);
    lcClobInsideV := f_border_text('insideV', p_table.border_inside_v);

    if 
        lcClobTop is not null or 
        lcClobBottom is not null or 
        lcClobLeft is not null or 
        lcClobRight is not null or
        lcClobInsideH is not null or
        lcClobInsideV is not null
        then
        p_add_clob_text('        <w:tblBorders>');
        p_add_clob_text(lcClobTop || lcClobBottom || lcClobLeft || lcClobRight || lcClobInsideH || lcClobInsideV);
        p_add_clob_text('        </w:tblBorders>');
    end if;

        
    p_add_clob_text('<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
    </w:tblPr>');

    --cells
    FOR v IN 1 .. p_table.rows_num LOOP
        p_add_clob_text('    <w:tr>');
        FOR s IN 1 .. p_table.columns_num LOOP
            if nvl(p_table.cells(v || ',' || s).merge_h, 0) <> -1 then
                p_add_clob_text('        <w:tc>');
                p_add_clob_text('        <w:tcPr>');
                
                --width 
                --for horizontal merge width is calculated as sum of merged cells width
                if nvl(p_table.cells(v || ',' || s).merge_h, 0) > 1 then
                    lnWidth := 0;
                    FOR t IN s .. (s + p_table.cells(v || ',' || s).merge_h - 1) LOOP
                        lnWidth := lnWidth + p_table.column_width(t);
                    END LOOP;
                else
                    lnWidth := p_table.column_width(s);
                end if;
                p_add_clob_text('            <w:tcW w:w="' || lnWidth || '" w:type="dxa"/>');
                
                --horizontal merge
                if nvl(p_table.cells(v || ',' || s).merge_h, 0) > 1 then
                    p_add_clob_text('        <w:gridSpan w:val="' || p_table.cells(v || ',' || s).merge_h || '"/>');
                end if;
                
                --vertical merge
                if p_table.cells(v || ',' || s).merge_v is not null then
                    p_add_clob_text('        ' || p_table.cells(v || ',' || s).merge_v);
                end if;
                
                --vertical alignment
                if p_table.cells(v || ',' || s).alignment_v <> 'TOP' then
                    p_add_clob_text('            <w:vAlign w:val="' || lower(p_table.cells(v || ',' || s).alignment_v) || '"/>');
                end if;
                
                --borders
                lcClobTop := f_border_text('top', p_table.cells(v || ',' || s).border_top);
                lcClobBottom := f_border_text('bottom', p_table.cells(v || ',' || s).border_bottom);
                lcClobLeft := f_border_text('left', p_table.cells(v || ',' || s).border_left);
                lcClobRight := f_border_text('right', p_table.cells(v || ',' || s).border_right);

                if lcClobTop is not null or lcClobBottom is not null or lcClobLeft is not null or lcClobRight is not null then
                    p_add_clob_text('        <w:tcBorders>');
                    p_add_clob_text(lcClobTop || lcClobBottom || lcClobLeft || lcClobRight);
                    p_add_clob_text('        </w:tcBorders>');
                end if;
                
                --background color
                if p_table.cells(v || ',' || s).background_color is not null then
                    p_add_clob_text('            <w:shd w:val="clear" w:color="auto" w:fill="' || lower(p_table.cells(v || ',' || s).background_color) || '"/>');
                end if;
                
                
                p_add_clob_text('        </w:tcPr>');
                
                --add paragraph/text
                p_xml_paragraph(p_doc_id, p_table.cells(v || ',' || s).paragraph);
                
                p_add_clob_text('        </w:tc>');
            end if;
        END LOOP;
        p_add_clob_text('    </w:tr>');
    END LOOP;

    --konec tabele
    p_add_clob_text('</w:tbl>');
END p_xml_table;

PROCEDURE p_xml_page_break(p_break r_break) IS
BEGIN
    p_add_clob_text('       <w:p>
        <w:r>
            <w:br w:type="page"/>
        </w:r>
    </w:p>');
END p_xml_page_break;

PROCEDURE p_xml_section_break(
    p_doc_id number,
    p_page r_page, 
    p_section_type varchar2,
    p_last boolean) IS
    
BEGIN
    --the last page at the end of document doesn't have w:p tag
    if not p_last then
        p_add_clob_text('<w:p>
			<w:pPr>');
    end if;
    
    --page data
	p_add_clob_text('<w:sectPr>' || 
            (CASE WHEN p_section_type <> 'nextPage' THEN '<w:type w:val="' || p_section_type || '"/> ' ELSE null END) ||
            (CASE WHEN p_page.header_ref is not null THEN '<w:headerReference w:type="default" r:id="rId' || grDoc(p_doc_id).containers(p_page.header_ref).rel_id || '"/>' ELSE null END) || 
            (CASE WHEN p_page.footer_ref is not null THEN '<w:footerReference w:type="default" r:id="rId' || grDoc(p_doc_id).containers(p_page.footer_ref).rel_id || '"/>' ELSE null END) || '
			<w:pgSz w:w="' || (CASE WHEN p_page.orientation <> 'landscape' THEN p_page.width ELSE p_page.height END) ||
			'" w:h="' || (CASE WHEN p_page.orientation <> 'landscape' THEN p_page.height ELSE p_page.width END) || '"' ||
			(CASE WHEN p_page.orientation = 'landscape' THEN ' w:orient="landscape"' ELSE null END) || '/>
			<w:pgMar w:top="' || p_page.margin_top || '" w:right="' || p_page.margin_right || '" w:bottom="' || p_page.margin_bottom || '" w:left="' || p_page.margin_left || '" w:header="' || p_page.header_h || '" w:footer="' || p_page.footer_h || '" w:gutter="0"/>
			<w:cols w:space="708"/>
			<w:docGrid w:linePitch="360"/>
		</w:sectPr>');

    if not p_last then
        p_add_clob_text('</w:pPr>
			</w:p>');
    end if;

END p_xml_section_break;





PROCEDURE p_container_content(
    p_doc_id number,
    p_container r_container,
    p_add_last_page boolean default true) IS
    
    lrPage r_page := grDoc(p_doc_id).default_page;
    lcSectionType varchar2(50) := 'nextPage';
    
BEGIN

	FOR t IN 1 .. p_container.elements.count LOOP
        
        if p_container.elements(t).element_type = 'PARAGRAPH' then
            p_xml_paragraph(p_doc_id, p_container.elements(t).paragraph);
            
        elsif p_container.elements(t).element_type = 'TABLE' then
            p_xml_table(p_doc_id, p_container.elements(t).table_data);
            
        elsif p_container.elements(t).element_type = 'BREAK' then
        
            if p_container.elements(t).break_data.break_type = 'PAGE' then
                p_xml_page_break(p_container.elements(t).break_data);
            elsif p_container.elements(t).break_data.break_type = 'SECTION' then
                p_xml_section_break(p_doc_id, lrPage, lcSectionType, false);
                lrPage := p_container.elements(t).break_data.page;
                lcSectionType := p_container.elements(t).break_data.section_type;
            end if;

        elsif p_container.elements(t).element_type = 'IMAGE' then
            p_xml_image(
                p_doc_id => p_doc_id,
                p_image_data => p_container.elements(t).image_data,
                p_paragraph_yn => 'Y'
            );
        end if;
        
	END LOOP;
	
	--last page
	if p_add_last_page then
        p_xml_section_break(p_doc_id, lrPage, lcSectionType, true);
    end if;

END;


FUNCTION f_document(p_doc_id number) RETURN clob IS
BEGIN
    --main clob variabe to null
    gcClob := null;
    
    --document start
    p_add_clob_text('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 wp14">
	<w:body>');

	--document content - first container is document
	p_container_content(p_doc_id, grDoc(p_doc_id).containers(1));
    
	--document end
	p_add_clob_text('	</w:body>
</w:document>');

    RETURN gcClob;
END;


FUNCTION f_header(
    p_doc_id number,
    p_header r_container) RETURN clob IS
BEGIN
    --main clob variabe to null
    gcClob := null;
    
    --header start
    p_add_clob_text('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14">');
    
    --header content
	p_container_content(p_doc_id, p_header, false);

    --header end
    p_add_clob_text('</w:hdr>');

    RETURN gcClob;
END;

FUNCTION f_make_document(
    p_doc_id number) RETURN blob IS
    
    lbWord blob;
    lcClob clob;
    lcClobHeaderRels clob;
    
BEGIN
    dbms_lob.createtemporary(lbWord, true);

    --main document files
    p_add_document_to_zip(lbWord, '[Content_Types].xml', f_content_types(p_doc_id));
    p_add_document_to_zip(lbWord, '_rels/.rels', f_rels);
    p_add_document_to_zip(lbWord, 'docProps/app.xml', f_app);
    p_add_document_to_zip(lbWord, 'docProps/core.xml', f_core(p_doc_id));
    p_add_document_to_zip(lbWord, 'word/_rels/document.xml.rels', f_document_xml_rels(p_doc_id) );
    p_add_document_to_zip(lbWord, 'word/theme/theme1.xml', f_theme);
    p_add_document_to_zip(lbWord, 'word/fontTable.xml', f_font_table);
    p_add_document_to_zip(lbWord, 'word/settings.xml', f_settings(p_doc_id));
    p_add_document_to_zip(lbWord, 'word/styles.xml', f_styles(p_doc_id));
    p_add_document_to_zip(lbWord, 'word/webSettings.xml', f_web_settings);
    p_add_document_to_zip(lbWord, 'word/document.xml', f_document(p_doc_id) );
    
    --additional documents, if used
    
    --numbering
    lcClob := f_numbering(p_doc_id);
    if lcClob is not null then
        p_add_document_to_zip(lbWord, 'word/numbering.xml', lcClob);
    end if;
    
    --headers and footers (relation XML documents too)
    if grDoc(p_doc_id).images.count > 0 then
        lcClobHeaderRels := f_container_xml_rels(p_doc_id);
    end if;
    
    FOR t IN 1 .. grDoc(p_doc_id).containers.count LOOP
        if grDoc(p_doc_id).containers(t).container_type = 'HEADER' then
            lcClob := f_header(p_doc_id, grDoc(p_doc_id).containers(t));
            p_add_document_to_zip(lbWord, 'word/header' || grDoc(p_doc_id).containers(t).rel_id || '.xml', lcClob);
            
            if grDoc(p_doc_id).images.count > 0 then
                p_add_document_to_zip(lbWord, 'word/_rels/header' || grDoc(p_doc_id).containers(t).rel_id || '.xml.rels', lcClobHeaderRels);
            end if;
        elsif grDoc(p_doc_id).containers(t).container_type = 'FOOTER' then
            lcClob := f_header(p_doc_id, grDoc(p_doc_id).containers(t));
            p_add_document_to_zip(lbWord, 'word/footer' || grDoc(p_doc_id).containers(t).rel_id || '.xml', lcClob);
            
            if grDoc(p_doc_id).images.count > 0 then
                p_add_document_to_zip(lbWord, 'word/_rels/footer' || grDoc(p_doc_id).containers(t).rel_id || '.xml.rels', lcClobHeaderRels);
            end if;
        end if;
    END LOOP;

    --images
    FOR t IN 1 .. grDoc(p_doc_id).images.count LOOP
        p_add_document_to_zip(lbWord, 'word/media/' || grDoc(p_doc_id).images(t).image_name, grDoc(p_doc_id).images(t).image_file);
    END LOOP;

    finish_zip(lbWord);

    RETURN lbWord;
    
END;


PROCEDURE p_save_file(
    p_document blob,
    p_file_name varchar2 default 'my_document.docx',
    p_folder varchar2 default 'MY_FOLDER'
    ) IS

    lfFile utl_file.file_type;
    lnLen pls_integer := 32767;
    
BEGIN
    lfFile := utl_file.fopen(p_folder, p_file_name, 'wb');
    FOR i in 0 .. trunc( (dbms_lob.getlength(p_document) - 1 ) / lnLen ) LOOP
        utl_file.put_raw(lfFile, dbms_lob.substr(p_document, lnLen, i * lnLen + 1));
    END LOOP;
    utl_file.fclose(lfFile);

END p_save_file;

PROCEDURE p_download_document(
    p_doc IN OUT blob,
    p_file_name varchar2,
    p_disposition varchar2 default 'attachment'  --values "attachment" and "inline"
    ) IS
BEGIN
    htp.init;
    OWA_UTIL.MIME_HEADER('application/pdf', FALSE);
    htp.p('Content-length: ' || dbms_lob.getlength(p_doc) ); 
    htp.p('Content-Disposition: ' || p_disposition || '; filename="' || p_file_name || '"' );
    OWA_UTIL.HTTP_HEADER_CLOSE;

    WPG_DOCLOAD.DOWNLOAD_FILE(p_doc);

    --free temporary lob IF it is temporary
    if dbms_lob.istemporary(p_doc) = 1 then
        DBMS_LOB.FREETEMPORARY(p_doc);
    end if;

    --uncomment only if You plan to download the generated document from the APEX
    --apex_application.stop_apex_engine;
END p_download_document;  

END zt_word;
/
