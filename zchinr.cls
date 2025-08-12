% File: zchinr.cls
% Copyright 2024-2025 Jasper Habicht (mail(at)jasperhabicht.de).
%
% This work may be distributed and/or modified under the
% conditions of the LaTeX Project Public License version 1.3c,
% available at http://www.latex-project.org/lppl/.
%
% This file is part of the `zchinr' package (The Work in LPPL)
% and all files in that bundle must be distributed together.
%
% This work has the LPPL maintenance status `maintained'.
%
\ProvidesExplClass{zchinr}
  {2025/08/12} {3.10} {Document class for ZChinR}

\bool_new:N \l_zchinr_prepress_bool
\bool_new:N \l_zchinr_standalone_bool

\keys_define:nn { zchinr / global } {
  prepress   .bool_set:N = \l_zchinr_prepress_bool ,
  prepress   .default:n  = { true } ,
  prepress   .initial:n  = { false } ,
  standalone .bool_set:N = \l_zchinr_standalone_bool ,
  standalone .default:n  = { true } ,
  standalone .initial:n  = { false }
}

\ProcessKeyOptions [ zchinr / global ]
\LoadClass [ a4paper , twoside , 10pt ] { report }

\AssignSocketPlug{build/column/outputbox}{floats-space-footnotes}

\msg_new:nnn { zchinr } { no-color-profile } {
  \msg_warning_text:n { zchinr } \iow_newline:
  Color ~ profile ~ not ~ defined ~ in ~ preamble.
}

\msg_new:nnn { zchinr } { pagination-unknown } {
  \msg_warning_text:n { zchinr } \iow_newline:
  Ignoring ~ unknown ~ pagination ~ setting: ~ #1.
}

\msg_new:nnn { zchinr } { header-style-unknown } {
  \msg_warning_text:n { zchinr } \iow_newline:
  Ignoring ~ unknown ~ header ~ style ~ setting: ~ #1.
}

\msg_new:nnn { zchinr } { toc-style-unknown } {
  \msg_warning_text:n { zchinr } \iow_newline:
  Ignoring ~ unknown ~ toc ~ style ~ setting: ~ #1.
}

\msg_new:nnn { zchinr } { cover-toc-style-unknown } {
  \msg_warning_text:n { zchinr } \iow_newline:
  Ignoring ~ unknown ~ cover ~ toc ~ style ~ setting: ~ #1.
}

\msg_new:nnn { zchinr } { wrong-page-count } {
  \msg_warning_text:n { zchinr } \iow_newline:
  Page ~ count ~ not ~ divisible ~ by ~ \int_use:N \l_zchinr_page_div_int : ~ #1.
}

\msg_new:nnn { zchinr } { missing-references } {
  \msg_warning_text:n { zchinr } \iow_newline:
  One ~ or ~ more ~ chapter ~ references ~ not ~ found. ~ Please ~ rerun ~ compilation.
}

\tl_new:N \l_zchinr_color_profile_name_tl
\tl_new:N \l_zchinr_color_profile_info_tl
\str_new:N \l_zchinr_color_profile_path_str

\int_new:N \l_zchinr_standalone_part_int

\int_new:N \g_zchinr_page_count_int
\int_new:N \l_zchinr_page_div_int

\tl_new:N \l_zchinr_issue_year_tl
\tl_new:N \l_zchinr_issue_volume_tl
\tl_new:N \l_zchinr_issue_number_tl
\tl_new:N \l_zchinr_issue_doi_tl
\tl_new:N \g_zchinr_issue_first_page_tl
\tl_new:N \g_zchinr_issue_last_page_tl
\bool_new:N \g__zchinr_issue_first_page_set_bool

\tl_new:N \l_zchinr_default_licence_tl
\tl_new:N \l_zchinr_default_licence_url_tl

\tl_new:N \l_zchinr_article_first_page_tl
\tl_new:N \l_zchinr_article_label_tl
\tl_new:N \l_zchinr_article_pages_tl
\tl_new:N \l_zchinr_article_pages_toc_tl
\tl_new:N \l_zchinr_article_doi_tl
\tl_new:N \l_zchinr_article_licence_tl
\tl_new:N \l_zchinr_article_licence_url_tl
\tl_new:N \l_zchinr_article_author_tl
\tl_new:N \l_zchinr_article_author_short_tl
\tl_new:N \l_zchinr_article_author_toc_tl
\tl_new:N \l_zchinr_article_author_copy_tl
\tl_new:N \l_zchinr_article_author_note_tl
\tl_new:N \l_zchinr_article_title_tl
\tl_new:N \l_zchinr_article_title_pdf_tl
\tl_new:N \l_zchinr_article_title_alt_tl
\tl_new:N \l_zchinr_article_title_short_tl
\tl_new:N \l_zchinr_article_title_toc_tl
\tl_new:N \l_zchinr_article_abstract_tl
\tl_new:N \l_zchinr_article_abstract_alt_tl
\tl_new:N \l_zchinr_article_category_tl
\int_new:N \l__zchinr_pages_from_int
\int_new:N \l__zchinr_pages_to_int
\str_new:N \l_zchinr_article_pagination_str
\str_new:N \l_zchinr_article_toc_style_str
\str_new:N \l_zchinr_article_cover_toc_style_str
\str_new:N \l_zchinr_article_header_style_str
\str_new:N \l_zchinr_article_page_style_str
\bool_new:N \l_zchinr_article_twocolumn_bool
\bool_new:N \l_zchinr_article_small_title_bool
\bool_new:N \l_zchinr_article_title_split_bool
\clist_new:N \l_zchinr_article_title_style_clist
\seq_new:N \g__zchinr_article_abstract_split_seq
\seq_new:N \g__zchinr_article_abstract_alt_split_seq

\dim_new:N \l_zchinr_article_abstract_width_dim
\dim_new:N \l_zchinr_title_distance_dim
\clist_new:N \l__zchinr_page_style_enlarge_clist

\box_new:N \l_zchinr_article_abstract_i_box
\box_new:N \l_zchinr_article_abstract_ii_box

\tl_new:N \l_zchinr_local_zchinr_tl
\tl_new:N \l_zchinr_local_abstract_tl
\tl_new:N \l_zchinr_local_doi_tl
\tl_new:N \l_zchinr_local_licence_tl
\tl_new:N \l_zchinr_local_page_tl

\tl_new:N \l__zchinr_header_mark_even_tl
\tl_new:N \l__zchinr_header_mark_odd_tl
\tl_new:N \l__zchinr_header_mark_page_tl

\tl_new:N \l__zchinr_documentation_header_tl
\tl_new:N \l__zchinr_toc_entry_tl
\bool_new:N \g__zchinr_missing_refs_bool

\tl_new:N \g__zchinr_cover_toc_tl

\tl_new:N \l__zchinr_fnmark_tl

\dim_new:N \l__zchinr_tabindent_dim
\dim_set:Nn \l__zchinr_tabindent_dim { \parindent }

\tl_new:N \l_zchinr_tmpa_tl
\tl_new:N \l_zchinr_tmpb_tl
\coffin_new:N \l_zchinr_tmpa_coffin

\iow_new:N \g__zchinr_xmp_iow

\cs_generate_variant:Nn \int_set:Nn { Ne }
\cs_generate_variant:Nn \int_to_roman:n { e }
\cs_generate_variant:Nn \property_record:nn { en }
\cs_generate_variant:Nn \property_ref:nn { en }
\prg_generate_conditional_variant:Nnn \int_compare:nNn { eNn } { T , TF }
\prg_generate_conditional_variant:Nnn \property_if_recorded:nn { en } { TF }

\keys_define:nn { zchinr / setup } {
  color ~ profile          .code:n      = { \keys_set:nn { zchinr / setup / color ~ profile } {#1} } ,
  color ~ profile / name   .tl_set:N    = \l_zchinr_color_profile_name_tl ,
  color ~ profile / name   .default:n   = { Custom } ,
  color ~ profile / info   .tl_set:N    = \l_zchinr_color_profile_info_tl ,
  color ~ profile / info   .default:n   = { none } ,
  color ~ profile / path   .str_set:N   = \l_zchinr_color_profile_path_str ,
  standalone ~ part        .int_set:N   = \l_zchinr_standalone_part_int ,
  standalone ~ part        .initial:n   = { 0 } ,
  title ~ distance         .dim_set:N   = \l_zchinr_title_distance_dim ,
  title ~ distance         .initial:n   = { 35mm } ,
  check ~ page ~ count     .int_set:N   = \l_zchinr_page_div_int ,
  check ~ page ~ count     .initial:n   = { 4 } ,
  localization             .code:n      = { \keys_set:nn { zchinr / setup / localization } {#1} } ,
  localization / zchinr    .tl_set:N    = \l_zchinr_local_zchinr_tl ,
  localization / zchinr    .initial:n   = { ZChinR } ,
  localization / abstract  .tl_set:N    = \l_zchinr_local_abstract_tl ,
  localization / abstract  .initial:n   = { Abstract } ,
  localization / doi       .tl_set:N    = \l_zchinr_local_doi_tl ,
  localization / doi       .initial:n   = { DOI } ,
  localization / licence   .tl_set:N    = \l_zchinr_local_licence_tl ,
  localization / licence   .initial:n   = { Lizenz } ,
  localization / page      .tl_set:N    = \l_zchinr_local_page_tl ,
  localization / page      .initial:n   = { S. } ,
  default / licence        .tl_set:N    = \l_zchinr_default_licence_tl ,
  default / licence        .initial:n   = { CC ~ BY ~ 4.0 } ,
  default / licence ~ url  .tl_set:N    = \l_zchinr_default_licence_url_tl ,
  default / licence ~ url  .initial:n   = { https://creativecommons.org/licenses/by/4.0/ } ,
  issue                    .code:n      = { \keys_set:nn { zchinr / setup / issue } {#1} } ,
  issue / year             .tl_set:N    = \l_zchinr_issue_year_tl ,
  issue / volume           .tl_set:N    = \l_zchinr_issue_volume_tl ,
  issue / number           .tl_set:N    = \l_zchinr_issue_number_tl ,
  issue / doi              .tl_set:N    = \l_zchinr_issue_doi_tl ,
  article                  .code:n      = {
    \keys_set_exclude_groups:nnn { zchinr / setup / article } { execute } {#1}
    \keys_set_groups:nnn { zchinr / setup / article } { execute } {#1}
  } ,
  article / pagination     .str_set:N   = \l_zchinr_article_pagination_str ,
  article / pagination     .initial:n   = { arabic } , % arabic , roman , none
  article / first ~ page   .tl_set:N    = \l_zchinr_article_first_page_tl ,
  article / first ~ page   .initial:n   = { } ,
  article / label          .tl_set:N    = \l_zchinr_article_label_tl ,
  article / label          .initial:n   = { } ,
  article / pages          .tl_set:N    = \l_zchinr_article_pages_tl ,
  article / pages          .initial:n   = {
    \c_colon_str \c_space_token \l_zchinr_local_page_tl \c_space_token
    \zchinr_pages:nn { \value{page} } { \__zchinr_current_chapter_last_page: }
    \int_compare:eNnT { \__zchinr_current_chapter_last_page: } < { 0 } {
      \bool_set_true:N \g__zchinr_missing_refs_bool
    }
  } ,
  article / pages ~ toc    .tl_set:N    = \l_zchinr_article_pages_toc_tl ,
  article / pages ~ toc    .initial:n   = {
    \zchinr_pages:nn { \thecontentspage } { \__zchinr_chapter_last_page:n { \value{zchinrchapterindex} } }
    \int_compare:eNnT { \__zchinr_chapter_last_page:n { \value{zchinrchapterindex} } } < { 0 } {
      \bool_set_true:N \g__zchinr_missing_refs_bool
    }
  } ,
  article / doi            .tl_set:N    = \l_zchinr_article_doi_tl ,
  article / licence        .tl_set:N    = \l_zchinr_article_licence_tl ,
  article / licence        .initial:n   = { \l_zchinr_default_licence_tl } ,
  article / licence ~ url  .tl_set:N    = \l_zchinr_article_licence_url_tl ,
  article / licence ~ url  .initial:n   = { l_zchinr_default_licence_url_tl } ,
  article / author         .tl_set:N    = \l_zchinr_article_author_tl ,
  article / author         .initial:n   = { } ,
  article / author ~ short .tl_set:N    = \l_zchinr_article_author_short_tl ,
  article / author ~ short .initial:n   = { \l_zchinr_article_author_tl } ,
  article / author ~ toc   .tl_set:N    = \l_zchinr_article_author_toc_tl ,
  article / author ~ toc   .initial:n   = { \l_zchinr_article_author_tl } ,
  article / author ~ copy  .tl_set:N    = \l_zchinr_article_author_copy_tl ,
  article / author ~ copy  .initial:n   = { \l_zchinr_article_author_tl } ,
  article / author ~ note  .tl_set:N    = \l_zchinr_article_author_note_tl ,
  article / author ~ note  .initial:n   = { } ,
  article / title          .tl_set:N    = \l_zchinr_article_title_tl ,
  article / title          .initial:n   = { } ,
  article / title ~ pdf    .tl_set:N    = \l_zchinr_article_title_pdf_tl ,
  article / title ~ pdf    .initial:n   = { \l_zchinr_article_title_tl } ,
  article / title ~ alt    .tl_set:N    = \l_zchinr_article_title_alt_tl ,
  article / title ~ alt    .initial:n   = { } ,
  article / title ~ short  .tl_set:N    = \l_zchinr_article_title_short_tl ,
  article / title ~ short  .initial:n   = { \l_zchinr_article_title_tl } ,
  article / title ~ toc    .tl_set:N    = \l_zchinr_article_title_toc_tl ,
  article / title ~ toc    .initial:n   = { \l_zchinr_article_title_tl } ,
  article / abstract       .tl_set:N    = \l_zchinr_article_abstract_tl ,
  article / abstract       .initial:n   = { } ,
  article / abstract ~ alt .tl_set:N    = \l_zchinr_article_abstract_alt_tl ,
  article / abstract ~ alt .initial:n   = { } ,
  article / category       .tl_set:N    = \l_zchinr_article_category_tl ,
  article / category       .initial:n   = { Aufsätze } ,
  article / two ~ columns  .bool_set:N  = \l_zchinr_article_twocolumn_bool ,
  article / two ~ columns  .default:n   = { true } ,
  article / two ~ columns  .initial:n   = { false } ,
  article / title ~ split
                           .bool_set:N   = \l_zchinr_article_title_split_bool ,
  article / title ~ split
                           .default:n   = { true } ,
  article / title ~ split
                           .initial:n   = { false } ,
  article / small ~ title  .bool_set:N  = \l_zchinr_article_small_title_bool ,
  article / small ~ title  .default:n   = { true } ,
  article / small ~ title  .initial:n   = { false } ,
  article / title ~ style  .clist_set:N = \l_zchinr_article_title_style_clist ,
  article / title ~ style  .initial:n   = { title } , % category , title , author , abstract , phantom
  article / header ~ style .str_set:N   = \l_zchinr_article_header_style_str ,
  article / header ~ style .initial:n   = { author ~ title } , % author title , title , category , contents
  article / toc ~ style    .str_set:N   = \l_zchinr_article_toc_style_str ,
  article / toc ~ style    .initial:n   = { author ~ first } , % author first , author last , no author , none
  article / cover ~ toc ~ style
                           .str_set:N   = \l_zchinr_article_cover_toc_style_str ,
  article / cover ~ toc ~ style
                           .initial:n   = { author ~ title } , % author title , title
  article / step ~ chapter .code:n      = { \__zchinr_step_chapter: } ,
  article / step ~ chapter .groups:n    = { execute } ,
  article / cover ~ toc    .code:n      = { \__zchinr_cover_toc_push:n {#1} } ,
  article / cover ~ toc    .groups:n    = { execute } ,
  article / page ~ style   .code:n      = {
    \str_set:Nn \l_zchinr_article_page_style_str {#1}
    \thispagestyle{zchinr #1}
  } ,
  article / page ~ style   .groups:n    = { execute } ,
  reset                    .code:n      = {
   \keys_set:nn { zchinr / setup / article } {
      pagination = { arabic } ,
      first ~ page = { } ,
      label = { } ,
      doi = { } ,
      licence = { \l_zchinr_default_licence_tl } ,
      licence ~ url = { \l_zchinr_default_licence_url_tl } ,
      author = { } ,
      author ~ short = { \l_zchinr_article_author_tl } ,
      author ~ toc = { \l_zchinr_article_author_tl } ,
      author ~ copy = { \l_zchinr_article_author_tl } ,
      author ~ note = { } ,
      title = { } ,
      title ~ pdf = { \l_zchinr_article_title_tl } ,
      title ~ alt = { } ,
      title ~ short = { \l_zchinr_article_title_tl } ,
      title ~ toc = { \l_zchinr_article_title_tl } ,
      abstract = { } ,
      abstract ~ alt = { } ,
      two ~ columns = { false } ,
      title ~ split = { false } ,
      small ~ title = { false } ,
      title ~ style = { title } ,
      header ~ style = { author ~ title } ,
      toc ~ style = { author ~ first }
    }
  }
}

\NewDocumentCommand { \zchinrsetup } { s m } {
  \IfBooleanF{#1}{
    \clearpage
  }
  \keys_set_exclude_groups:nnn { zchinr / setup } { execute } {#2}
  \keys_set_groups:nnn { zchinr / setup } { execute } {#2}
}

\hook_new_pair:nn { zchinr / titleblock / before } { zchinr / titleblock / after }
\hook_new_pair:nn { zchinr / category / before } { zchinr / category / after }
\hook_new_pair:nn { zchinr / title / before } { zchinr / title / after }
\hook_new_pair:nn { zchinr / subcontents / before } { zchinr / subcontents / after }
\hook_new_pair:nn { zchinr / abstract / before } { zchinr / abstract / after }
\hook_new:n { zchinr / abstract / split }

\cs_new_protected:Npn \zchinr_halfquad: {
  \hskip .5em \relax
}

\cs_new_protected:Npn \zchinr_hfillrule: {
  \leavevmode\leaders\hrule height .75pt\hfill\kern\z@
}

\cs_new_protected:Npn \zchinr_pages:nn #1#2 {
  \int_set:Ne \l__zchinr_pages_from_int {#1}
  \int_set:Ne \l__zchinr_pages_to_int {#2}
  \int_compare:nNnTF { \l__zchinr_pages_from_int } = { \l__zchinr_pages_to_int } {
    \int_compare:nNnTF { \l__zchinr_pages_from_int } < { 0 }
      { ? } { \int_use:N \l__zchinr_pages_from_int }
  } {
    \int_compare:nNnTF { \l__zchinr_pages_from_int } < { 0 }
      { ? } { \int_use:N \l__zchinr_pages_from_int }
    --
    \int_compare:nNnTF { \l__zchinr_pages_from_int } < { 0 }
      { ? } { \int_use:N \l__zchinr_pages_to_int }
  }
}

\RequirePackage[english, ngerman]{babel}

\RequirePackage{fontspec}
\setmainfont{TeX ~ Gyre ~ Pagella}
\setsansfont{TeX ~ Gyre ~ Heros}
\newfontfamily\zchinrlatindisplay{TeX ~ Gyre ~ Heros ~ Cn}

\sys_if_engine_luatex:T {
  \RequirePackage{luatexja-fontspec}
  \setmainjfont[ItalicFont={Noto ~ Serif ~ CJK ~ SC}]
    {Noto ~ Serif ~ CJK ~ SC}
  \setsansjfont[ItalicFont={Noto ~ Serif ~ CJK ~ SC}]
    {Noto ~ Sans ~ CJK ~ SC}

  % prevent German quotation marks from being spaced
  \ltjsetparameter{
    alxspmode = {"201E, preonly},
    alxspmode = {"201C, postonly}
  }
}

\sys_if_engine_xetex:T {
  \RequirePackage{xeCJK}
  \setCJKmainfont[ItalicFont={Noto ~ Serif ~ CJK ~ SC}]
    {Noto ~ Serif ~ CJK ~ SC}
  \setCJKsansfont[ItalicFont={Noto ~ Serif ~ CJK ~ SC}]
    {Noto ~ Sans ~ CJK ~ SC}

  % prevent German quotation marks from being spaced
  \xeCJKDeclareCharClass{NormalSpace}{"201E, "201A, "002C}
  \xeCJKDeclareCharClass{NormalSpace}{"201C, "2018, "0060}
}

\RequirePackage{unicode-math}
\setmathfont{TeX ~ Gyre ~ Pagella ~ Math}

\NewDocumentCommand{\zhs}{ m }{\textup{#1}}
\NewDocumentCommand{\zht}{ m }{\textup{#1}}

\RequirePackage{microtype}

\RequirePackage[
  inner=25mm,
  outer=30mm,
  top=30mm,
  bottom=25mm,
  headsep=10mm,
  footskip=5mm
]{geometry}
\setlength{\columnsep}{7.5mm}

\RequirePackage{l3draw}

\cs_if_exist:NF \draw_set_linewidth:n {
  \cs_set_eq:NN \draw_set_linewidth:n \draw_linewidth:n
}

\cs_if_exist:NF \draw_set_baseline:n {
  \cs_set_eq:NN \draw_set_baseline:n \draw_baseline:n
}

\cs_if_exist:NF \draw_set_evenodd_rule: {
  \cs_set_eq:NN \draw_set_evenodd_rule: \draw_evenodd_rule:
}

\cs_new_protected:Npn \zchinr_cropmark_nw: {
  \draw_begin:
    \bool_set_false:N \l_draw_bb_update_bool
    \draw_set_linewidth:n { .25pt }
    \draw_transform_shift:n { 10mm , -10mm }
    \draw_path_moveto:n { -5pt , 0pt }
    \draw_path_lineto:n { -20pt , 0pt }
    \draw_path_moveto:n { 0pt , 5pt }
    \draw_path_lineto:n { 0pt , 20pt }
    \draw_path_use:n { stroke }
  \draw_end:
}
\cs_new_protected:Npn \zchinr_cropmark_ne: {
  \draw_begin:
    \bool_set_false:N \l_draw_bb_update_bool
    \draw_set_linewidth:n { .25pt }
    \draw_transform_shift:n { \paperwidth - 10mm , -10mm }
    \draw_path_moveto:n { 5pt , 0pt }
    \draw_path_lineto:n { 20pt , 0pt }
    \draw_path_moveto:n { 0pt , 5pt }
    \draw_path_lineto:n { 0pt , 20pt }
    \draw_path_use:n { stroke }
  \draw_end:
}
\cs_new_protected:Npn \zchinr_cropmark_sw: {
  \draw_begin:
    \bool_set_false:N \l_draw_bb_update_bool
    \draw_set_linewidth:n { .25pt }
    \draw_transform_shift:n { 10mm , 10mm - \paperheight }
    \draw_path_moveto:n { -5pt , 0pt }
    \draw_path_lineto:n { -20pt , 0pt }
    \draw_path_moveto:n { 0pt , -5pt }
    \draw_path_lineto:n { 0pt , -20pt }
    \draw_path_use:n { stroke }
  \draw_end:
}
\cs_new_protected:Npn \zchinr_cropmark_se: {
  \draw_begin:
    \bool_set_false:N \l_draw_bb_update_bool
    \draw_set_linewidth:n { .25pt }
    \draw_transform_shift:n { \paperwidth - 10mm , 10mm - \paperheight }
    \draw_path_moveto:n { 5pt , 0pt }
    \draw_path_lineto:n { 20pt , 0pt }
    \draw_path_moveto:n { 0pt , -5pt }
    \draw_path_lineto:n { 0pt , -20pt }
    \draw_path_use:n { stroke }
  \draw_end:
}

\bool_if:NTF \l_zchinr_prepress_bool {
  \color_set:nnn { zchinrblue } { cmyk } { 1 , .5 , 0 , .5 }
} {
  \color_set:nnn { zchinrblue } { rgb } { 0 , .25 , .5 }
}

\bool_if:NT \l_zchinr_prepress_bool {
  \geometry{
    layout=a4paper,
    paperwidth=230mm,
    paperheight=317mm,
    layouthoffset=10mm,
    layoutvoffset=10mm
  }

  \sys_if_engine_luatex:TF {
    \pdfvariable ~ pagesattr {
      /MediaBox ~ [0.0000 ~ 0.0000 ~ 651.9690 ~ 898.5830]
    }
  } {
    \special{pdf:put ~ @pages  ~ <<
      /MediaBox ~ [0.0000 ~ 0.0000 ~ 651.9690 ~ 898.5830]
    >>}
  }

  \AddToHook{shipout/foreground}{
    \sys_if_engine_luatex:TF {
      \pdfvariable ~ pageattr {
        /BleedBox ~ [19.8425 ~ 19.8425 ~ 632.1265 ~ 878.7405]
        /TrimBox ~ [28.3465 ~ 28.3465 ~ 623.6225 ~ 870.2365]
      }
    } {
      \special{pdf:put ~ @thispage ~ <<
        /BleedBox ~ [19.8425 ~ 19.8425 ~ 632.1265 ~ 878.7405]
        /TrimBox ~ [28.3465 ~ 28.3465 ~ 623.6225 ~ 870.2365]
      >>}
    }
    \zchinr_cropmark_nw:
    \zchinr_cropmark_ne:
    \zchinr_cropmark_sw:
    \zchinr_cropmark_se:
  }

  \AddToHook{begindocument}{
    \sys_if_engine_luatex:TF {
      \directlua{
        pdf.setminorversion(3)
      }
      \pdfextension ~ info {
        /GTS_PDFXVersion ~ (PDF/X-1:2001)
        /GTS_PDFXConformance ~ (PDF/X-1a:2001)
      }
    } {
      \special{pdf:minorversion ~ 3}
      \special{pdf:docinfo ~ <<
        /GTS_PDFXVersion ~ (PDF/X-1:2001)
        /GTS_PDFXConformance ~ (PDF/X-1a:2001)
      >>}
    }
    \str_if_empty:NTF \l_zchinr_color_profile_path_str {
      \msg_warning:nn { zchinr } { no-color-profile }
      \sys_if_engine_luatex:TF {
        \pdfextension ~ catalog {
          /PageMode ~ /UseNone
          /OutputIntents ~ [<<
            /Type ~ /OutputIntent
            /S ~ /GTS_PDFX
            /OutputConditionIdentifier ~ (Custom)
            /Info ~ (none)
            /RegistryName ~ (http://www.color.org/)
          >>]
        }
      } {
        \special{pdf:put ~ @catalog ~ <<
          /PageMode ~ /UseNone
          /OutputIntents ~ [<<
            /Type ~ /OutputIntent
            /S ~ /GTS_PDFX
            /OutputConditionIdentifier ~ (Custom)
            /Info ~ (none)
            /RegistryName ~ (http://www.color.org/)
          >>]
        >>}
      }
    } {
      \sys_if_engine_luatex:TF {
        \immediate \pdfextension obj ~ stream ~
          attr{/N ~ 4 ^^J /Alternate/DeviceCMYK} ~
          file{\l_zchinr_color_profile_path_str}
        \pdfextension ~ catalog {
          /PageMode ~ /UseNone
          /OutputIntents ~ [<<
            /Type ~ /OutputIntent
            /S ~ /GTS_PDFX
            /DestOutputProfile ~ \the\numexpr\pdffeedback ~ lastobj\relax\space 0 ~ R
            /OutputConditionIdentifier ~ (\l_zchinr_color_profile_name_tl)
            /Info ~ (\l_zchinr_color_profile_info_tl)
            /RegistryName ~ (http://www.color.org/)
          >>]
        }
      } {
        \special{pdf:fstream ~ @cmykdata ~ (\l_zchinr_color_profile_path_str) ~ <<
          /N ~ 4
          /Alternate/DeviceCMYK
        >>}
        \special{pdf:put ~ @catalog ~ <<
          /PageMode ~ /UseNone
          /OutputIntents ~ [<<
            /Type ~ /OutputIntent
            /S ~ /GTS_PDFX
            /DestOutputProfile ~ @cmykdata
            /OutputConditionIdentifier ~ (\l_zchinr_color_profile_name_tl)
            /Info ~ (\l_zchinr_color_profile_info_tl)
            /RegistryName ~ (http://www.color.org/)
          >>]
        >>}
      }
    }
  }
}

\RequirePackage{array}
\RequirePackage{longtable}
\RequirePackage{booktabs}
\RequirePackage{enumitem}
\RequirePackage{titletoc}
\RequirePackage[pagestyles]{titlesec}

\frenchspacing
\clubpenalty=10000
\widowpenalty=10000
\doublehyphendemerits=5000000

\setlength{\parindent}{15pt}
\setlength{\parskip}{2pt plus 8pt minus 1pt}

\setlength{\arrayrulewidth}{.75pt}

\setlength{\LTpre}{0pt}
\setlength{\LTpost}{0pt}

% Style figure and table captions.
\RequirePackage[
  labelfont=bf,
  labelsep=period,
]{caption}

\cs_new:Npn \__zchinr_generate_xmp: {
  \iow_open:Nn \g__zchinr_xmp_iow { \c_sys_jobname_str .xmp }
  \iow_now:Ne \g__zchinr_xmp_iow {
    <?xpacket ~ begin='' ~ id=''?> \iow_newline:
    <x:xmpmeta ~ xmlns:x='adobe:ns:meta/'> \iow_newline:
    <rdf:RDF ~ xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns \c_hash_str "> \iow_newline:
    <rdf:Description ~ rdf:about="" \iow_newline:
      xmlns:dc="http://purl.org/dc/elements/1.1/" \iow_newline:
      xmlns:xmpRights="http://ns.adobe.com/xap/1.0/rights/"> \iow_newline:    <dc:rights> \iow_newline:
    <rdf:Alt> \iow_newline:
    <rdf:li ~ xml:lang="x-default">Copyright ~ (c) ~ ZChinR ~ \l_zchinr_issue_year_tl ; ~
      \l_zchinr_local_licence_tl \c_colon_str ~ \l_zchinr_default_licence_tl </rdf:li> \iow_newline:
    </rdf:Alt> \iow_newline:
    </dc:rights> \iow_newline:
    <xmpRights:Marked>True</xmpRights:Marked> \iow_newline:
    <xmpRights:Certificate> \l_zchinr_default_licence_url_tl </xmpRights:Certificate> \iow_newline:
    <xmpRights:WebStatement> \l_zchinr_default_licence_url_tl </xmpRights:WebStatement> \iow_newline:
    </rdf:Description> \iow_newline:
    </rdf:RDF> \iow_newline:
    </x:xmpmeta> \iow_newline:
    <?xpacket ~ end='r'?>
  }
  \iow_close:N \g__zchinr_xmp_iow
}

% Set up hyperreferences
\PassOptionsToPackage{hyphens}{url}
\tl_if_exist:NF \pdfcreationdate {
  \tl_new:N \pdfcreationdate
  \sys_if_engine_luatex:TF {
    \tl_set:Nn \pdfcreationdate {
      \directlua{
        tex.sprint(os.date("D:
          \c_percent_str Y
          \c_percent_str m
          \c_percent_str d
          \c_percent_str H
          \c_percent_str M
          \c_percent_str S
          Z"))
      }
    }
  } {
    \tl_set_eq:NN \pdfcreationdate \creationdate
  }
}
\bool_if:NTF \l_zchinr_prepress_bool {
  \RequirePackage[bookmarks=false]{hyperref}
  \AddToHook{begindocument}{
    \NoHyper
  }
} {
  \RequirePackage[hidelinks]{hyperref}
  \sys_if_engine_luatex:T {
    \AddToHook{begindocument}{
      \__zchinr_generate_xmp:
      \group_begin:
        \pdfcompresslevel=0
        \immediate \pdfobj stream ~
          attr{/Type ~ /Metadata ~ /Subtype ~ /XML} ~
          file{ \c_sys_jobname_str .xmp}
        \pdfextension ~ catalog {
          /Metadata ~ \the \pdflastobj \space 0 ~ R
        }
      \group_end:
    }
  }
}
\bool_if:NTF \l_zchinr_standalone_bool {
  \property_if_recorded:nnT { standalone } { zchinr / chapter-title } {
    \AddToHook{begindocument}{
      \hypersetup{
        pdftitle={\property_ref:nn { standalone } { zchinr / chapter-title }}
      }
    }
  }
  \property_if_recorded:nnT { standalone } { zchinr / chapter-author } {
    \AddToHook{begindocument}{
      \hypersetup{
        pdfauthor={\property_ref:nn { standalone } { zchinr / chapter-author }}
      }
    }
  }
  \AddToHook{begindocument}{
    \hypersetup{
      pdfmoddate={\pdfcreationdate},
      pdftrapped={False}
    }
  }
} {
  \AddToHook{begindocument}{
    \hypersetup{
      pdftitle={\l_zchinr_local_zchinr_tl \c_space_token \l_zchinr_issue_year_tl},
      pdfmoddate={\pdfcreationdate},
      pdftrapped={False}
    }
  }
}

% Style header and footer. Header stays empty while footer will contain journal information and page number.
\cs_new_protected:Npn \__zchinr_sethead_default: {
  \sethead
    [\l__zchinr_header_mark_page_tl]
    [\l__zchinr_header_mark_odd_tl]
    [\l_zchinr_local_zchinr_tl \c_space_token \l_zchinr_issue_year_tl]
    {\l_zchinr_local_zchinr_tl \c_space_token \l_zchinr_issue_year_tl}
    {\l__zchinr_header_mark_even_tl}
    {\l__zchinr_header_mark_page_tl}
}

\cs_new_protected:Npn \__zchinr_setfoot_contents: {
  \setfoot
    {{\parbox[b]{0.4\textwidth}{\raggedright
      \l_zchinr_local_doi_tl \c_colon_str \c_space_token
        \href{https://doi.org/\l_zchinr_issue_doi_tl}{\l_zchinr_issue_doi_tl}
    }}}
    {}
    {{\parbox[b]{0.4\textwidth}{\raggedleft
      \copyright \c_space_token \l_zchinr_issue_year_tl \c_space_token
        \l_zchinr_local_zchinr_tl \\
    }}}
}

\cs_new_protected:Npn \__zchinr_setfoot_title: {
  \setfoot
    {{\parbox[b]{0.4\textwidth}{\raggedright
      ZChinR \c_space_token \l_zchinr_issue_year_tl
        \l_zchinr_article_pages_tl \\
      \l_zchinr_local_doi_tl \c_colon_str \c_space_token
        \href{https://doi.org/\l_zchinr_article_doi_tl}{\l_zchinr_article_doi_tl}
    }}}
    {}
    {{\parbox[b]{0.4\textwidth}{\raggedleft
      \copyright \c_space_token \l_zchinr_issue_year_tl \c_space_token
        \l_zchinr_article_author_copy_tl \\
      \l_zchinr_local_licence_tl \c_colon_str \c_space_token
        \href{\l_zchinr_article_licence_url_tl}{\l_zchinr_article_licence_tl}
    }}}
}

\cs_new_protected:Npn \__zchinr_sethead_empty: {
  \sethead{}{}{}
}

\cs_new_protected:Npn \__zchinr_setfoot_empty: {
  \setfoot{}{}{}
}

\newpagestyle{zchinrdefault}{
  \__zchinr_sethead_default:
  \__zchinr_setfoot_empty:
}
\pagestyle{zchinrdefault}

\newpagestyle{zchinrdefaulttitle}{
  \__zchinr_sethead_default:
  \__zchinr_setfoot_title:
}

\newpagestyle{zchinremptytitle}{
  \__zchinr_sethead_empty:
  \__zchinr_setfoot_title:
}

\newpagestyle{zchinremptycontents}{
  \__zchinr_sethead_empty:
  \__zchinr_setfoot_contents:
}

\newpagestyle{zchinrempty}{
  \__zchinr_sethead_empty:
  \__zchinr_setfoot_empty:
}

\clist_map_inline:nn { defaulttitle , emptytitle , emptycontents } {
  \clist_put_right:Ne \l__zchinr_page_style_enlarge_clist {
    \tl_to_str:n {#1}
  }
}

\newcounter{zchinrchapterindex}

% Style section titles.
\titleformat{\section}{\fontsize{11}{13}\selectfont\bfseries}{}{0pt}{}
\titleformat{\subsection}{\fontsize{11}{13}\selectfont\bfseries}{}{0pt}{}
\titleformat{\subsubsection}{\fontsize{11}{13}\selectfont\bfseries}{}{0pt}{}
\titlespacing*{\section}{0pt}{5pt plus 10pt}{5pt}
\titlespacing*{\subsection}{0pt}{5pt plus 10pt}{5pt}
\titlespacing*{\subsubsection}{0pt}{5pt plus 10pt}{5pt}

\titleformat{\paragraph}[runin]{\bfseries}{}{2em}{\noindent}
\titleformat{\subparagraph}[runin]{\bfseries}{}{2em}{\noindent}
\titlespacing*{\paragraph}{0pt}{5pt plus 10pt}{1em plus 5pt minus 10pt}
\titlespacing*{\subparagraph}{0pt}{5pt plus 10pt}{1em plus 5pt minus 10pt}

\titlecontents{part}[0pt]{\bigskip\Large\scshape\bfseries}{}{}{}
\titlecontents{chapter}[15pt]{\vskip 5pt\contentsmargin{5em}}{}{}{
  \hfill \l_zchinr_article_pages_toc_tl \hspace*{-5em}
}

\titlecontents{section}[0pt]{}{\hspace*{-15pt}}{}
  { \zchinr_halfquad: \titlerule*[5pt]{.}\contentspage}
\titlecontents{subsection}[15pt]{}{\hspace*{-15pt}}{}
  { \zchinr_halfquad: \titlerule*[5pt]{.}\contentspage}
\titlecontents{subsubsection}[30pt]{}{\hspace*{-15pt}}{}
  { \zchinr_halfquad: \titlerule*[5pt]{.}\contentspage}
\titlecontents{paragraph}[45pt]{}{\hspace*{-15pt}}{}
  { \zchinr_halfquad: \titlerule*[5pt]{.}\contentspage}
\titlecontents{subparagraph}[60pt]{}{\hspace*{-15pt}}{}
  { \zchinr_halfquad: \titlerule*[5pt]{.}\contentspage}

% Style footnotes.
\tl_set:Nn \l__zchinr_fnmark_tl {
  \@thefnmark
}
\cs_set:Npn \@makefntext #1 {
	\dim_set:Nn \leftskip { 2em }
	\skip_horizontal:n { -2em }
	\hbox_to_wd:nn { 2em } { \l__zchinr_fnmark_tl }
	#1
}
\tl_set:Nn \footnoterule {}
\skip\footins 10pt plus 15pt minus 5pt

% Style URLs.
\urlstyle{same}
\DeclareUrlCommand\Hurl{\def\UrlLeft{<\,}\def\UrlRight{\,>}}
\tl_put_right:Nn \UrlBreaks {\do\^\do\_}

% Create command for automatically linked e-mail addresses.
\NewDocumentCommand { \email } { m } {
  \href{mailto:#1}{\protect\nolinkurl{#1}}
}

\NewDocumentCommand { \currentissueyear } { } {
  \tl_use:N \l_zchinr_issue_year_tl
}

\NewDocumentCommand { \currentissuevolume } { } {
  \tl_use:N \l_zchinr_issue_volume_tl
}

\NewDocumentCommand { \currentissuenumber } { } {
  \tl_use:N \l_zchinr_issue_number_tl
}

\NewDocumentCommand { \currentissuefirstpage } { } {
  \property_if_recorded:nnTF { issue } { zchinr / issue-first-page } {
    \property_ref:nn { issue } { zchinr / issue-first-page }
  } {
    \tl_use:N \g_zchinr_issue_first_page_tl
  }
}

\NewDocumentCommand { \currentissuelastpage } { } {
  \property_if_recorded:nnTF { issue } { zchinr / issue-last-page } {
    \property_ref:nn { issue } { zchinr / issue-last-page }
  } {
    \tl_use:N \g_zchinr_issue_first_page_tl
  }
}

% Create documentation environment.
\NewDocumentEnvironment { documentation } { } {
  \tl_set:Nn \arraystretch { 1.25 }
  \noindent\begin{longtable}{
    @{}
      >{ \dim_set:Nn \l__zchinr_tabindent_dim { 2em }
        \hskip \l__zchinr_tabindent_dim }
      p{54mm}
    @{\hskip\columnsep}
      >{ \dim_set:Nn \l__zchinr_tabindent_dim { 15pt }
        \hskip \l__zchinr_tabindent_dim }
      p{\dimexpr\textwidth - 54mm - \columnsep}
    @{}
  }
} {
  \end{longtable}
}

\NewDocumentCommand { \documentationheader } { O{c} O{} m o o } {
  \leavevmode
  \IfNoValueF{#4}{
    \phantomsection
    \addcontentsline{toc}{#4}{\IfNoValueTF{#5}{#3}{#5}}
  }
  \hskip -\l__zchinr_tabindent_dim
  \rule{0pt}{\dimexpr\baselineskip+5pt}
  \str_if_eq:nnT {#1} { c } { \centering }
  \str_if_eq:nnT {#1} { l } { \raggedright }
  \str_if_eq:nnT {#1} { r } { \raggedleft }
  \str_if_eq:nnF {#2} { n } { \bfseries }
  #3 \arraybackslash
  \rule[-10pt]{0pt}{0pt}
}

\NewDocumentCommand { \documentationtocentry } { m m } {
  \leftskip\dimexpr #1\l__zchinr_tabindent_dim
  \hskip\dimexpr -2\l__zchinr_tabindent_dim \relax #2
}

\NewDocumentCommand { \zchinrbreakline } { } {
  \unskip\parfillskip 0pt\relax
}

\NewDocumentCommand { \zchinrcontinueline } { } {
  \hskip\dimexpr -1\l__zchinr_tabindent_dim \relax
  \parfillskip 0pt plus 1fil\relax
}

\NewDocumentCommand { \zchinrbreakpage } { } {
  \zchinrbreakline
  \pagebreak
  \zchinrcontinueline
}

% Create addresses environment.
\NewDocumentEnvironment { addresses } { } {
  \tl_set:Nn \arraystretch { 1.75 }
  \noindent\begin{longtable}{
    @{}
      p{\dimexpr.5\textwidth - .5\columnsep}
    @{\hskip\columnsep}
      p{\dimexpr.5\textwidth - .5\columnsep}
    @{}
  }
} {
  \end{longtable}
  \bigskip
}

\NewDocumentCommand { \addressessection } { s m o } {
  \IfBooleanT{#1}{
    \phantomsection
    \addcontentsline{toc}{chapter}{\IfNoValueTF{#2}{#1}{#2}}
  }
  \vskip 10pt plus 5pt minus 2.5pt\relax
  \noindent{\Large #2}\par
}

\NewDocumentCommand { \addressesheader } { s m o o } {
  \leavevmode
  \IfBooleanT{#1}{
    \IfNoValueF{#3}{
      \phantomsection
      \addcontentsline{toc}{#3}{\IfNoValueTF{#4}{#2}{#4}}
   }
  }
  \rule{0pt}{\dimexpr\baselineskip+15pt}
  \raggedright\bfseries\arraybackslash #2
}

\NewExpandableDocumentCommand { \addressesspan } { m } {
  \multicolumn{2}{@{}p{\textwidth}@{}}{#1}
}

\NewDocumentCommand { \addressessep } { m } {
  \char"00B7{}
}

% Create imprint environment.
\NewDocumentEnvironment { imprint } { } {
  \tl_set:Nn \arraystretch { 1.5 }
  \noindent\begin{longtable}{
    @{}
      >{\raggedright}
      p{32.5mm}
    @{\hskip\columnsep}
      >{\raggedright}
      p{\dimexpr\textwidth - 32.5mm - 27.5mm - 1.5\columnsep}
    @{\hskip .5\columnsep}
      >{\raggedleft\let\newline\\\arraybackslash}
      p{27.5mm}
    @{}
  }
} {
  \end{longtable}
}

\NewExpandableDocumentCommand { \imprintspan } { m } {
  \multicolumn{2}{@{}p{\dimexpr\textwidth - 32.5mm - \columnsep}@{}}{\small #1}
}

\NewExpandableDocumentCommand { \imprintline } { } {
  \multicolumn{2}{@{}p{\dimexpr\textwidth - 32.5mm - \columnsep}@{}}{
    \zchinr_hfillrule:
  }
}

% Create commands for TOC.
\NewDocumentCommand { \tocentry } { m }{
  \bigskip
  \noindent{\Large\scshape\bfseries \text_lowercase:n {#1} }
}

\NewDocumentCommand { \printissuetoc } { } {
  \startcontents
  \printcontents{}{-1}[0]{}
}

% Redefine abstract environment.
\RenewDocumentEnvironment { abstract } { O{\textwidth} } {
  \noindent
  \begin{minipage}[t]{#1}
  \setlength{\parindent}{15pt}
  \noindent
} {
  \end{minipage}
}

% Create subcontents environment.
\NewDocumentEnvironment { subcontents } { } {
  \noindent
  \list{}{\leftmargin=27.5mm \rightmargin=27.5mm}
  \item\relax
} {
  \endlist
}

\cs_new_protected:Npn \__zchinr_cover_toc_push:n #1 {
  \tl_set:Ne \l_zchinr_tmpa_tl { \l_zchinr_article_author_toc_tl }
  \tl_set:Ne \l_zchinr_tmpb_tl { \l_zchinr_article_title_toc_tl }
  \tl_if_blank:nTF {#1} {
    \str_case_e:nnF { \l_zchinr_article_cover_toc_style_str } {
      { title } {
        \tl_gput_right:Ne \g__zchinr_cover_toc_tl {
          \exp_not:N \protect \exp_not:N \item{} \l_zchinr_tmpb_tl
        }
      }
      { author ~ title } {
        \tl_gput_right:Ne \g__zchinr_cover_toc_tl {
          \exp_not:N \protect \exp_not:N \item{} \exp_not:N \emph{ \l_zchinr_tmpa_tl , }
          \  \l_zchinr_tmpb_tl
        }
      }
    } {
      \msg_warning:nnn { zchinr } { cover-toc-style-unknown }
        { \l_zchinr_article_cover_toc_style_str }
    }
  } {
    \tl_gput_right:Ne \g__zchinr_cover_toc_tl { \exp_not:N \protect \exp_not:N \item{} #1 }
  }
}

% The following automatically set the "first page" and "last page" of the issue.
% The "first page" is the first page with arabic page numbers.
% The "last page" is the last page with arabic page numbers.
\property_new:nnnn { zchinr / cover-toc } { shipout } { } { \exp_not:o { \g__zchinr_cover_toc_tl } }
\property_new:nnnn { zchinr / issue-first-page } { now } { ? } { \tl_use:N \g_zchinr_issue_first_page_tl }
\property_new:nnnn { zchinr / issue-last-page } { now } { ? } { \tl_use:N \g_zchinr_issue_last_page_tl }

\property_new:nnnn { zchinr / chapter-last-page } { now } { ? } { \int_eval:n { \value{page} - 1 } }
\property_new:nnnn { zchinr / chapter-author } { now } { } { \tl_use:N \l_zchinr_article_author_tl }
\property_new:nnnn { zchinr / chapter-title } { now } { } { \tl_use:N \l_zchinr_article_title_pdf_tl }

% Set the "first page" to the current page number if arabic, and set it only once.
% After each page break, set the "last page" to the current page number if arabic.
\AddToHook{shipout/before}{
  \str_if_eq:VnTF \l_zchinr_article_pagination_str { arabic } {
    \bool_if:NF \g__zchinr_issue_first_page_set_bool {
      \tl_gset:Ne \g_zchinr_issue_first_page_tl { \thepage }
      \bool_gset_true:N \g__zchinr_issue_first_page_set_bool
    }
    \tl_gset:Ne \g_zchinr_issue_last_page_tl { \thepage }
  }
}

\AddToHook{shipout/after}{
  \int_gincr:N \g_zchinr_page_count_int
}

% At \end{document}, write "first page" and "last page" to .aux file.
\AddToHook{enddocument/afterlastpage}{
  \group_begin:
    \cs_set:Npn \write { \tex_immediate:D \tex_write:D }
    \property_record:nn { issue }
      { zchinr / cover-toc , zchinr / issue-first-page , zchinr / issue-last-page }
    \bool_if:NTF \l_zchinr_standalone_bool {
      \property_gset:nnnn { zchinr / chapter-last-page } { now } { ? } { \int_eval:n { \value{page} } }
    } {
      \property_gset:nnnn { zchinr / chapter-last-page } { now } { ? } { \int_eval:n { \value{page} - 1 } }
    }
    \property_record:en { chapter : \int_eval:n { \value{chapter} } } { zchinr / chapter-last-page }
    \bool_if:NT \l_zchinr_standalone_bool {
      \property_record:nn { standalone } { zchinr / chapter-author , zchinr / chapter-title }
    }
  \group_end:

  \bool_if:NT \g__zchinr_missing_refs_bool {
    \msg_warning:nn { zchinr } { missing-references }
  }
  \fp_compare:nNnF {
    \g_zchinr_page_count_int - \l_zchinr_page_div_int *
    floor ( \g_zchinr_page_count_int / \l_zchinr_page_div_int )
  } = { 0 } {
    \bool_if:NF \l_zchinr_standalone_bool {
      \msg_warning:nnV { zchinr } { wrong-page-count }
        \g_zchinr_page_count_int
    }
  }
}

\cs_new_protected:Npn \__zchinr_step_chapter: {
  \stepcounter{chapter}
  \int_compare:eNnT { \value{chapter} } > { 1 } {
    \addtocontents{toc}{\protect\setcounter{zchinrchapterindex}{ \int_to_arabic:n { \value{chapter} } }}
    \property_record:en { chapter : \int_eval:n { \value{chapter} - 1 } } { zchinr / chapter-last-page }
  }
}

\cs_new:Npn \__zchinr_chapter_last_page:n #1 {
  \property_if_recorded:enTF { chapter : \int_eval:n {#1} } { zchinr / chapter-last-page } {
    \property_ref:en { chapter : \int_eval:n {#1} } { zchinr / chapter-last-page }
  } {
    -1
  }
}

\NewDocumentCommand { \chapterlastpage } { m } {
  \int_compare:eNnTF { \__zchinr_chapter_last_page:n {#1} } < { 0 } {
    \bool_set_true:N \g__zchinr_missing_refs_bool
    \textbf{?}
  } {
    \__zchinr_chapter_last_page:n {#1}
  }
}

\cs_new:Npn \__zchinr_current_chapter_last_page: {
  \property_if_recorded:enTF { chapter : \int_eval:n { \value{chapter} } } { zchinr / chapter-last-page } {
    \property_ref:en { chapter : \int_eval:n { \value{chapter} } } { zchinr / chapter-last-page }
  } {
    -1
  }
}

\NewDocumentCommand { \currentchapterlastpage } { } {
  \int_compare:eNnTF { \__zchinr_current_chapter_last_page: } < { 0 } {
    \bool_set_true:N \g__zchinr_missing_refs_bool
    \textbf{?}
  } {
    \__zchinr_current_chapter_last_page:
  }
}

\NewDocumentCommand { \printtitle } { } {
  % Using the page counter is save here, as the page title is typically placed at the beginning of the page.
  \tl_if_empty:NTF \l_zchinr_article_first_page_tl {
    \tl_set:Ne \l_zchinr_article_first_page_tl { \value{page} }
  } {
    \setcounter{page}{ \l_zchinr_article_first_page_tl }
  }
  \setcounter{footnote}{0}
  \startcontents[zchinrsub]

  \str_case_e:nnF { \l_zchinr_article_pagination_str } {
    { none } {
      \tl_set:Nn \l__zchinr_header_mark_page_tl { }
    }
    { roman } {
      \tl_set:Nn \l__zchinr_header_mark_page_tl { \roman{page} }
    }
    { arabic } {
      \tl_set:Nn \l__zchinr_header_mark_page_tl { \arabic{page} }
    }
  } {
    \msg_warning:nnn { zchinr } { pagination-unknown }
      { \l_zchinr_article_pagination_str }
  }

  \str_case_e:nnF { \l_zchinr_article_header_style_str } {
    { category } {
      \tl_set_eq:NN \l__zchinr_header_mark_even_tl \l_zchinr_article_category_tl
      \tl_set_eq:NN \l__zchinr_header_mark_odd_tl \l_zchinr_article_category_tl
      \clist_if_in:NnTF \l_zchinr_article_title_style_clist { category } {
        \tl_if_empty:NF \l_zchinr_article_doi_tl {
          \str_set:Nn \l_zchinr_article_page_style_str { emptytitle }
          \thispagestyle{zchinremptytitle}
        }
      } {
        \tl_if_empty:NF \l_zchinr_article_doi_tl {
          \str_set:Nn \l_zchinr_article_page_style_str { defaulttitle }
          \thispagestyle{zchinrdefaulttitle}
        }
      }
    }
    { title } {
      \tl_set_eq:NN \l__zchinr_header_mark_even_tl \l_zchinr_article_title_short_tl
      \tl_set_eq:NN \l__zchinr_header_mark_odd_tl \l_zchinr_article_title_short_tl
      \tl_if_empty:NF \l_zchinr_article_doi_tl {
        \str_set:Nn \l_zchinr_article_page_style_str { emptytitle }
        \thispagestyle{zchinremptytitle}
      }
    }
    { author ~ title } {
      \tl_set:Nn \l__zchinr_header_mark_even_tl {
        { \itshape \l_zchinr_article_author_short_tl , } ~ \l_zchinr_article_title_short_tl
      }
      \tl_set:Nn \l__zchinr_header_mark_odd_tl {
        { \itshape \l_zchinr_article_author_short_tl , } ~ \l_zchinr_article_title_short_tl
      }
      \tl_if_empty:NF \l_zchinr_article_doi_tl {
        \str_set:Nn \l_zchinr_article_page_style_str { emptytitle }
        \thispagestyle{zchinremptytitle}
      }
    }
    { contents } {
      \tl_set_eq:NN \l__zchinr_header_mark_even_tl \l_zchinr_article_category_tl
      \tl_set_eq:NN \l__zchinr_header_mark_odd_tl \l_zchinr_article_category_tl
      \tl_if_empty:NF \l_zchinr_issue_doi_tl {
        \str_set:Nn \l_zchinr_article_page_style_str { emptycontents }
        \thispagestyle{zchinremptycontents}
      }
    }
  } {
    \tl_if_empty:NF \l_zchinr_article_doi_tl {
      \str_set:Nn \l_zchinr_article_page_style_str { defaulttitle }
      \thispagestyle{zchinrdefaulttitle}
    }
    \msg_warning:nnn { zchinr } { header-style-unknown }
      { \l_zchinr_article_header_style_str }
  }

  \tl_clear:N \l__zchinr_toc_entry_tl
  \str_case_e:nnF { \l_zchinr_article_toc_style_str } {
    { author ~ first } {
      \tl_set:Nn \l__zchinr_toc_entry_tl {
        { \itshape \l_zchinr_article_author_toc_tl , } ~
          \l_zchinr_article_title_toc_tl
      }
    }
    { author ~ last } {
      \tl_set:Nn \l__zchinr_toc_entry_tl {
        \l_zchinr_article_title_toc_tl \newline
          { \itshape ( \l_zchinr_article_author_toc_tl ) }
      }
    }
    { no ~ author } {
      \tl_set:Nn \l__zchinr_toc_entry_tl {
        \l_zchinr_article_title_toc_tl
      }
    }
    { none } { }
  } {
    \msg_warning:nnn { zchinr } { toc-style-unknown }
      { \l__zchinr_toc_entry_tl }
  }

  \tl_set:Nn \thefootnote { \fnsymbol{footnote} }
  \tl_set:Nn \l__zchinr_fnmark_tl {
    \textsuperscript{\@thefnmark}
  }
  \hook_use:n { zchinr / titleblock / before }
  \__zchinr_article_title_block:
  \tl_if_empty:NF \l_zchinr_article_author_note_tl {
    \footnotetext{\l_zchinr_article_author_note_tl}
    \addtocounter{footnote}{-1}
  }
  \hook_use:n { zchinr / abstract / split }
  \tl_set:Nn \thefootnote { \arabic{footnote} }
  \tl_set:Nn \l__zchinr_fnmark_tl {
    \@thefnmark
  }
  \hook_use:n { zchinr / titleblock / after }
  \@afterindentfalse\@afterheading
}

\cs_new_protected:Npn \__zchinr_article_author: {
  \l_zchinr_article_author_tl
  \tl_if_empty:NF \l_zchinr_article_author_note_tl {
    \, \footnotemark
  }
}

\cs_new_protected:Npn \__zchinr_article_category: {
  \bool_if:NF \l_zchinr_standalone_bool {
    \tl_if_empty:NF \l__zchinr_toc_entry_tl {
      \phantomsection
      \addcontentsline{toc}{part}{
        \texorpdfstring{ \text_lowercase:n { \l_zchinr_article_category_tl } }
          {\l_zchinr_article_category_tl}
      }
    }
  }
  \noindent
  {
    % Using the page counter is save here, as the category is typically placed at the beginning of the page.
    \int_if_even:nTF { \value{page} } {
      \raggedright
    } {
      \raggedleft
    }
    \LARGE\scshape\bfseries \text_lowercase:n { \l_zchinr_article_category_tl }
    \zchinr_halfquad: \zchinr_hfillrule: \par
  }
  \vskip 7mm\relax
}

\cs_new_protected:Npn \__zchinr_article_title_large: {
  \bool_if:NF \l_zchinr_standalone_bool {
    \__zchinr_step_chapter:
    \tl_if_empty:NF \l__zchinr_toc_entry_tl {
      \phantomsection
      \addcontentsline{toc}{chapter}{
        \texorpdfstring{\l__zchinr_toc_entry_tl}{\l_zchinr_article_title_pdf_tl}
      }
    }
  }
  \noindent\parbox[t][\l_zchinr_title_distance_dim][t]{\textwidth}{
    \raggedright
    {\huge\l_zchinr_article_title_tl\par}
    \clist_if_in:NnT \l_zchinr_article_title_style_clist { author } {
      \bigskip
      {\fontsize{11}{13}\selectfont\itshape \__zchinr_article_author: \par}
    }
  }
  \tl_if_empty:NF \l_zchinr_article_label_tl {
    \label { chapter: \l_zchinr_article_label_tl }
  }
  \vskip 4mm\relax
  \bigskip
}

\cs_new_protected:Npn \__zchinr_article_title_small: {
  \bool_if:NF \l_zchinr_standalone_bool {
    \__zchinr_step_chapter:
    \tl_if_empty:NF \l__zchinr_toc_entry_tl {
      \phantomsection
      \addcontentsline{toc}{chapter}{
        \texorpdfstring{\l__zchinr_toc_entry_tl}{\l_zchinr_article_title_pdf_tl}
      }
    }
  }
  \noindent
  {\fontsize{11}{13}\selectfont\bfseries\l_zchinr_article_title_tl\par}
  \tl_if_empty:NF \l_zchinr_article_label_tl {
    \label { chapter: \l_zchinr_article_label_tl }
  }
  \bigskip\nopagebreak
  \noindent
  {\itshape \__zchinr_article_author: }\par
  \bigskip\nopagebreak
}

\cs_new_protected:Npn \__zchinr_article_subcontents: {
  \begin{subcontents}
    \printcontents[zchinrsub]{}{1}{}
  \end{subcontents}
  \par
  \vskip 12mm\relax
}

\NewDocumentCommand { \zchinrsplit } { } { \relax }

\cs_generate_variant:Nn \seq_gset_split:Nnn { Nno }

\cs_new_protected:Npn \__zchinr_article_abstract_title: {
  \noindent
  {\fontsize{11}{13}\selectfont\itshape\bfseries\l_zchinr_local_abstract_tl\par}
  \vskip 5pt plus 5pt\relax
}

\cs_new_protected:Npn \__zchinr_article_abstract: {
  \box_clear:N \l_zchinr_article_abstract_i_box
  \box_clear:N \l_zchinr_article_abstract_ii_box
  \bool_if:NT \l_zchinr_article_title_split_bool {
    \seq_gset_split:Nno \g__zchinr_article_abstract_split_seq
      { \zchinrsplit } { \l_zchinr_article_abstract_tl \null }
    \tl_set:No \l_zchinr_article_abstract_tl {
      \seq_item:Nn \g__zchinr_article_abstract_split_seq { 1 }
    }
    \tl_if_empty:NF \l_zchinr_article_abstract_alt_tl {
      \seq_gset_split:Nno \g__zchinr_article_abstract_alt_split_seq
        { \zchinrsplit } { \l_zchinr_article_abstract_alt_tl }
      \tl_set:No \l_zchinr_article_abstract_alt_tl {
        \seq_item:Nn \g__zchinr_article_abstract_alt_split_seq { 1 }
      }
    }
  }
  \tl_if_empty:NTF \l_zchinr_article_abstract_alt_tl {
    \bool_if:NTF \l_zchinr_article_title_split_bool {
      \tl_if_empty:eF {
        \seq_item:Nn \g__zchinr_article_abstract_split_seq { 1 }
      } {
        \vbox_set:Nn \l_zchinr_article_abstract_i_box {
          \__zchinr_article_abstract_title:
          \begin{abstract}
            \itshape
            \l_zchinr_article_abstract_tl \linebreak
          \end{abstract}
          \vspace*{-1\baselineskip}
        }
      }
      \tl_if_empty:eF {
        \seq_item:Nn \g__zchinr_article_abstract_split_seq { 2 }
      } {
        \vbox_set:Nn \l_zchinr_article_abstract_ii_box {
          \tl_if_empty:eT {
            \l_zchinr_article_abstract_tl
          } {
            \__zchinr_article_abstract_title:
          }
          \begin{abstract}
            \itshape
            \seq_item:Nn \g__zchinr_article_abstract_split_seq { 2 }
          \end{abstract}
          \vskip 6mm\relax
          \noindent
          \zchinr_hfillrule:
          \par
          \vskip 8mm minus 5pt\relax
        }
      }
    } {
      \vbox_set:Nn \l_zchinr_article_abstract_i_box {
        \__zchinr_article_abstract_title:
        \begin{abstract}
          \itshape
          \l_zchinr_article_abstract_tl
        \end{abstract}
        \vskip 6mm\relax
        \noindent
        \zchinr_hfillrule:
        \par
        \vskip 8mm minus 5pt\relax
      }
    }
  } {
    \bool_if:NTF \l_zchinr_article_title_split_bool {
      \tl_if_empty:eF {
        \seq_item:Nn \g__zchinr_article_abstract_split_seq { 1 }
      } {
        \vbox_set:Nn \l_zchinr_article_abstract_i_box {
          \__zchinr_article_abstract_title:
          \begin{abstract}[ \l_zchinr_article_abstract_width_dim ]
            \itshape
            \l_zchinr_article_abstract_tl \linebreak
          \end{abstract}
          \hskip\columnsep
          \begin{abstract}[ \l_zchinr_article_abstract_width_dim ]
            \itshape{\bfseries\l_zchinr_article_title_alt_tl} ~ --- ~
            \l_zchinr_article_abstract_alt_tl \linebreak
          \end{abstract}
          \vspace*{-1\baselineskip}
        }
      }
      \tl_if_empty:eF {
        \seq_item:Nn \g__zchinr_article_abstract_split_seq { 2 }
      } {
        \vbox_set:Nn \l_zchinr_article_abstract_ii_box {
          \tl_if_empty:eT {
            \seq_item:Nn \g__zchinr_article_abstract_split_seq { 1 }
          } {
            \__zchinr_article_abstract_title:
          }
          \begin{abstract}[ \l_zchinr_article_abstract_width_dim ]
            \itshape
            \seq_item:Nn \g__zchinr_article_abstract_split_seq { 2 }
          \end{abstract}
          \hskip\columnsep
          \begin{abstract}[ \l_zchinr_article_abstract_width_dim ]
            \itshape
            \tl_if_empty:eT {
              \seq_item:Nn \g__zchinr_article_abstract_split_seq { 1 }
            } {
              \itshape{\bfseries\l_zchinr_article_title_alt_tl} ~ --- ~
            }
            \seq_item:Nn \g__zchinr_article_abstract_alt_split_seq { 2 }
          \end{abstract}
          \vskip 6mm\relax
          \noindent
          \zchinr_hfillrule:
          \par
          \vskip 8mm minus 5pt\relax
        }
      }
    } {
      \vbox_set:Nn \l_zchinr_article_abstract_i_box {
        \__zchinr_article_abstract_title:
        \begin{abstract}[ \l_zchinr_article_abstract_width_dim ]
          \itshape
          \l_zchinr_article_abstract_tl
        \end{abstract}
        \hskip\columnsep
        \begin{abstract}[ \l_zchinr_article_abstract_width_dim ]
          \itshape{\bfseries\l_zchinr_article_title_alt_tl} ~ --- ~
          \l_zchinr_article_abstract_alt_tl
        \end{abstract}
        \vskip 6mm\relax
        \noindent
        \zchinr_hfillrule:
        \par
        \vskip 8mm minus 5pt\relax
      }
    }
  }
  \bool_if:NTF \l_zchinr_article_twocolumn_bool {
    \bool_if:NTF \l_zchinr_article_title_split_bool {
      \box_use:N \l_zchinr_article_abstract_i_box
      \AddToHookNext{zchinr/abstract/split}{
        \clearpage
        \twocolumn[{
          \box_use:N \l_zchinr_article_abstract_ii_box
        }]
        \sloppy\flushbottom
      }
    } {
      \twocolumn[{
        \box_use:N \l_zchinr_article_abstract_i_box
      }]
      \sloppy\flushbottom
    }
  } {
    \bool_if:NTF \l_zchinr_article_title_split_bool {
      \box_use:N \l_zchinr_article_abstract_i_box
      \AddToHookNext{zchinr/abstract/split}{
        \clearpage
        \box_use:N \l_zchinr_article_abstract_ii_box
      }
    } {
      \box_use:N \l_zchinr_article_abstract_i_box
    }
  }
}

\NewDocumentCommand { \printabstract } { } {
  \__zchinr_article_abstract:
}

\cs_new_protected:Npn \__zchinr_article_title: {
  \clist_if_in:NnT \l_zchinr_article_title_style_clist { category } {
    \hook_use:n { zchinr / category / before }
    \__zchinr_article_category:
    \hook_use:n { zchinr / category / after }
  }
  \bool_if:NTF \l_zchinr_article_small_title_bool {
    \clist_if_in:NnT \l_zchinr_article_title_style_clist { title } {
      \hook_use:n { zchinr / title / before }
      \__zchinr_article_title_small:
      \hook_use:n { zchinr / title / after }
    }
  } {
    \clist_if_in:NnT \l_zchinr_article_title_style_clist { title } {
      \hook_use:n { zchinr / title / before }
      \__zchinr_article_title_large:
      \hook_use:n { zchinr / title / after }
    }
    \clist_if_in:NnT \l_zchinr_article_title_style_clist { contents } {
      \hook_use:n { zchinr / subcontents / before }
      \__zchinr_article_subcontents:
      \hook_use:n { zchinr / subcontents / after }
    }
    \clist_if_in:NnT \l_zchinr_article_title_style_clist { abstract } {
      \hook_use:n { zchinr / abstract / before }
      \__zchinr_article_abstract:
      \hook_use:n { zchinr / abstract / after }
    }
  }
}

\IfFormatAtLeastTF{2025/06/01}{}{
  \NewHook{build/column/before}
  \let\orig@makecol\@makecol
  \def\@makecol{
    \UseHook{build/column/before}
    \orig@makecol
  }
}

\cs_new_protected:Npn \__zchinr_article_title_block: {
  \tl_if_in:NnTF \l_zchinr_article_abstract_tl { \zchinrsplit } {
    \bool_set_true:N \l_zchinr_article_title_split_bool
  } {
    \bool_set_false:N \l_zchinr_article_title_split_bool
  }
  \dim_set:Nn \l_zchinr_article_abstract_width_dim {
    ( \textwidth - \columnsep ) / 2
  }
  \bool_lazy_and:nnTF {
    \bool_if_p:N \l_zchinr_article_twocolumn_bool
  } {
    \bool_if_p:n {
      \bool_not_p:n { \l_zchinr_article_title_split_bool }
    }
  } {
    \twocolumn[
      \__zchinr_article_title:
    ]
    \sloppy\flushbottom
    \clist_if_in:NVT \l__zchinr_page_style_enlarge_clist \l_zchinr_article_page_style_str {
      \enlargethispage{-10mm}
      \AddToHookNext{build/column/before}{\enlargethispage{-10mm}}
    }
  } {
    \onecolumn
    \__zchinr_article_title:
    \clist_if_in:NVT \l__zchinr_page_style_enlarge_clist \l_zchinr_article_page_style_str {
      \enlargethispage{-10mm}
    }
  }
  \clist_if_in:NnT \l_zchinr_article_title_style_clist { phantom } {
    \__zchinr_step_chapter:
    \tl_if_empty:NF \l__zchinr_toc_entry_tl {
      \phantomsection
      \addcontentsline{toc}{chapter}{
        \texorpdfstring{\l__zchinr_toc_entry_tl}{\l_zchinr_article_title_pdf_tl}
      }
    }
  }
}

\NewDocumentCommand { \inputpart } { O{0} m } {
  \bool_if:NTF \l_zchinr_standalone_bool {
    \int_compare:nNnT { \l_zchinr_standalone_part_int } = {#1} {
      \input{#2}
    }
  } {
    \input{#2}
  }
}

\newlist{covertoc}{itemize}{1}
\setlist[covertoc]{label={}, leftmargin=0pt, nosep, itemsep=10pt, before=\raggedright}

\NewDocumentCommand { \printcovertoc } { } {
  \property_if_recorded:nnT { issue } { zchinr / cover-toc } {
    \begin{covertoc}
      \property_ref:nn { issue } { zchinr / cover-toc }
    \end{covertoc}
  }
}

\NewDocumentEnvironment { drawabsolute } { s O{ 0cm , 0cm } +b } {
  \AddToHookNext{shipout/background}{
    \nullfont
    \draw_begin:
      \bool_set_false:N \l_draw_bb_update_bool
      \bool_if:NT \l_zchinr_prepress_bool {
        \draw_transform_shift:n { #2 }
      }
      #3
    \draw_end:
  }
  \IfBooleanF {#1} {
    \str_set:Nn \l_zchinr_article_page_style_str { empty }
    \thispagestyle{zchinrempty}
    \null
    \clearpage
  }
} { }

\NewDocumentCommand { \drawrectangle } { o m m } {
  \draw_scope_begin:
    \IfValueT{#1}{
      \color_select:n {#1}
    }
    \draw_path_rectangle:nn {#2} {#3}
    \draw_path_use:n { fill }
  \draw_scope_end:
}

\NewDocumentCommand { \drawcoffin } { o m m O{hc} O{vc} +m } {
  \draw_scope_begin:
    \IfValueT{#1}{
      \color_select:n {#1}
    }
    \vcoffin_set:Nnn \l_zchinr_tmpa_coffin {#3} { #6 \par }
    \draw_transform_shift:n {#2}
    \draw_coffin_use:Nnn \l_zchinr_tmpa_coffin {#4} {#5}
  \draw_scope_end:
}

\NewDocumentCommand { \drawinput } { o m m } {
  \draw_scope_begin:
    \IfValueT{#1}{
      \color_select:n {#1}
    }
    \draw_transform_shift:n {#2}
    \input{#3}
  \draw_scope_end:
}

\NewDocumentCommand { \oasymbol } { } {
  \draw_begin:
    \draw_set_baseline:n { -.175em }
    \draw_scope_begin:
      \draw_set_evenodd_rule:
      \draw_path_circle:nn { 0 , 0 } { 1 }
      \draw_path_use_clear:n { fill }
      \draw_path_circle:nn { 0 , 0 } { 3 }
      \draw_path_circle:nn { 0 , 0 } { 2 }
      \draw_path_use_clear:n { fill }
      \draw_path_moveto:n { 2.75 , 0 }
      \draw_path_lineto:n { 2.75 , 3.5 }
      \draw_path_arc:nnn { 0 } { 180 } { 2.75 }
      \draw_path_lineto:n { -1.75 , 3.5 }
      \draw_path_arc:nnn { 180 } { 0 } { 1.75 }
      \draw_path_lineto:n { 1.75 , 2 }
      \draw_path_close:
      \draw_path_use_clear:n { fill }
    \draw_scope_end:
  \draw_end:
}

\NewDocumentCommand { \ccsymbol } { O{1} } {
  \draw_begin:
    \draw_transform_scale:n { 0.1 * #1 }
    \draw_set_baseline:n { -.375em }
    \draw_scope_begin:
      \draw_set_evenodd_rule:
      \draw_path_circle:nn { 0 , 0 } { 50 }
      \draw_path_circle:nn { 0 , 0 } { 40 }
      \draw_path_use_clear:n { fill }
      \draw_scope_begin:
        \draw_transform_shift:n { -15 , 0 }
        \draw_path_moveto:n { \draw_point_polar:nnn { 7 } { 9 } { 30 } }
        \draw_path_arc:nnnn { 30 } { 330 } { 7 } { 9 }
        \draw_path_lineto:n { \draw_point_polar:nnn { 17 } { 18 } { 330 } }
        \draw_path_arc:nnnn { 330 } { 30 } { 17 } { 18 }
        \draw_path_close:
        \draw_path_use_clear:n { fill }
      \draw_scope_end:
      \draw_scope_begin:
        \draw_transform_shift:n { 17.5 , 0 }
        \draw_path_moveto:n { \draw_point_polar:nnn { 7 } { 9 } { 30 } }
        \draw_path_arc:nnnn { 30 } { 330 } { 7 } { 9 }
        \draw_path_lineto:n { \draw_point_polar:nnn { 17 } { 18 } { 330 } }
        \draw_path_arc:nnnn { 330 } { 30 } { 17 } { 18 }
        \draw_path_close:
        \draw_path_use_clear:n { fill }
      \draw_scope_end:
    \draw_scope_end:
  \draw_end:
}

\NewDocumentCommand { \ccbysymbol } { O{1} } {
  \draw_begin:
    \draw_transform_scale:n { 0.1 * #1 }
    \draw_set_baseline:n { -.375em }
    \draw_scope_begin:
      \draw_set_evenodd_rule:
      \draw_path_circle:nn { 0 , 0 } { 50 }
      \draw_path_circle:nn { 0 , 0 } { 40 }
      \draw_path_use:n { fill }
      \draw_path_moveto:n { -9 , -32.5 }
      \draw_path_lineto:n { 9 , -32.5 }
      \draw_path_lineto:n { 9 , -10 }
      \draw_path_lineto:n { 15 , -10 }
      \draw_path_lineto:n { 15 , 10 }
      \draw_path_arc:nnn { 0 } { 90 } { 5 }
      \draw_path_lineto:n { -10 , 15 }
      \draw_path_arc:nnn { 90 } { 180 } { 5 }
      \draw_path_lineto:n { -15 , -10 }
      \draw_path_lineto:n { -9 , -10 }
      \draw_path_close:
      \draw_path_circle:nn { 0 , 25 } { 7.5 }
      \draw_path_use_clear:n { fill }
    \draw_scope_end:
  \draw_end:
}

% EOF
