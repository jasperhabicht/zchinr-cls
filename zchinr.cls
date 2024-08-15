\ProvidesExplClass{zchinr}
  {2024/08/14} {2.0} {Document class for ZChinR}

\bool_new:N \l__zchinr_prepress_bool
\bool_new:N \l__zchinr_standalone_bool

\keys_define:nn { zchinr / global } { 
  prepress   .bool_set:N = \l__zchinr_prepress_bool ,
  prepress   .default:n  = { true } ,
  prepress   .initial:n  = { false } ,
  standalone .bool_set:N = \l__zchinr_standalone_bool ,
  standalone .default:n  = { true } ,
  standalone .initial:n  = { false } 
}

\ProcessKeyOptions [ zchinr / global ]
\LoadClass [ a4paper , twoside , 10pt ] { report }

\bool_new:N \l_zchinr_pagination_bool

\tl_new:N \l_zchinr_color_profile_id_tl
\tl_new:N \l_zchinr_color_profile_info_tl
\tl_new:N \l_zchinr_color_profile_path_tl

\tl_new:N \l_zchinr_issue_year_tl
\tl_new:N \l_zchinr_issue_number_tl
\tl_new:N \l_zchinr_issue_issn_tl

\tl_new:N \l_zchinr_article_doi_tl
\tl_new:N \l_zchinr_article_licence_tl
\tl_new:N \l_zchinr_article_licence_url_tl
\tl_new:N \l_zchinr_article_author_tl
\tl_new:N \l_zchinr_article_author_short_tl
\tl_new:N \l_zchinr_article_author_copy_tl
\tl_new:N \l_zchinr_article_author_note_tl
\tl_new:N \l_zchinr_article_title_tl
\tl_new:N \l_zchinr_article_title_pdf_tl
\tl_new:N \l_zchinr_article_title_alt_tl
\tl_new:N \l_zchinr_article_title_short_tl
\tl_new:N \l_zchinr_article_abstract_tl
\tl_new:N \l_zchinr_article_abstract_alt_tl
\tl_new:N \l_zchinr_article_category_tl
\int_new:N \l_zchinr_article_page_first_int
\int_new:N \l_zchinr_article_page_last_int
\bool_new:N \l_zchinr_article_twocolumn_bool
\clist_new:N \l_zchinr_article_title_style_clist

\dim_new:N \l_zchinr_title_distance_dim

\tl_new:N \l_zchinr_local_zchir_tl
\tl_new:N \l_zchinr_local_abstract_tl
\tl_new:N \l_zchinr_local_issn_tl
\tl_new:N \l_zchinr_local_doi_tl
\tl_new:N \l_zchinr_local_licence_tl

\keys_define:nn { zchir / color profile } {
  id   .tl_set:N   = \l_zchinr_color_profile_id_tl ,
  id   .default:n  = { Custom } ,
  info .tl_set:N   = \l_zchinr_color_profile_info_tl ,
  info .default:n  = { none } ,
  path .tl_set:N   = \l_zchinr_color_profile_path_tl 
}

\NewDocumentCommand { \zchinrcolorprofile } { m } {
  \keys_set:nn { zchinr / color profile } {#1}
}

\keys_define:nn { zchinr / setup } { 
  title ~ distance         .dim_set:N   = \l_zchinr_title_distance_dim ,
  title ~ distance         .initial:n   = { 35mm } ,
  pagination               .bool_set:N  = \l_zchinr_pagination_bool ,
  pagination               .default:n   = { true } ,
  pagination               .initial:n   = { true } ,
  localization             .code:n      = { \keys_set:nn { zchinr / setup / localization } {#1} } ,
  localization / zchinr    .tl_set:N    = \l_zchinr_local_zchinr_tl ,
  localization / zchinr    .initial:n   = { ZChinR } ,
  localization / abstract  .tl_set:N    = \l_zchinr_local_abstract_tl ,
  localization / abstract  .initial:n   = { Abstract } ,
  localization / issn      .tl_set:N    = \l_zchinr_local_issn_tl ,
  localization / issn      .initial:n   = { ISSN } ,
  localization / doi       .tl_set:N    = \l_zchinr_local_doi_tl ,
  localization / doi       .initial:n   = { DOI } ,
  localization / licence   .tl_set:N    = \l_zchinr_local_licence_tl ,
  localization / licence   .initial:n   = { Lizenz } ,
  issue                    .code:n      = { \keys_set:nn { zchinr / setup / issue } {#1} } ,
  issue / year             .tl_set:N    = \l_zchinr_issue_year_tl , 
  issue / number           .tl_set:N    = \l_zchinr_issue_number_tl , 
  issue / issn             .tl_set:N    = \l_zchinr_issue_issn_tl ,
  issue / issn             .initial:n   = { 1613-5768, ~ online ~ 2366-7125 } ,
  article                  .code:n      = { \keys_set:nn { zchinr / setup / article } {#1} } ,
  article / doi            .tl_set:N    = \l_zchinr_article_doi_tl , 
  article / licence        .tl_set:N    = \l_zchinr_article_licence_tl ,
  article / licence        .initial:n   = { CC ~ BY ~ 4.0 } ,
  article / licence ~ url  .tl_set:N    = \l_zchinr_article_licence_url_tl ,
  article / licence ~ url  .initial:n   = { https://creativecommons.org/licenses/by/4.0/ } ,
  article / pages          .code:n      = { \__zchinr_pages_split:n {#1} } , 
  article / page ~ first   .int_set:N   = \l_zchinr_page_first_int , 
  article / page ~ first   .initial:n   = { 1 } , 
  article / page ~ last    .int_set:N   = \l_zchinr_page_last_int , 
  article / page ~ last    .initial:n   = { 1 } , 
  article / author         .tl_set:N    = \l_zchinr_article_author_tl , 
  article / author ~ short .tl_set:N    = \l_zchinr_article_author_short_tl , 
  article / author ~ short .initial:n   = { \l_zchinr_article_author_tl } , 
  article / author ~ copy  .tl_set:N    = \l_zchinr_article_author_copy_tl , 
  article / author ~ copy  .initial:n   = { \l_zchinr_article_author_tl } , 
  article / author ~ note  .tl_set:N    = \l_zchinr_article_author_note_tl , 
  article / title          .tl_set:N    = \l_zchinr_article_title_tl , 
  article / title ~ pdf    .tl_set:N    = \l_zchinr_article_title_pdf_tl , 
  article / title ~ pdf    .initial:n   = { \l_zchinr_article_title_tl } ,
  article / title ~ alt    .tl_set:N    = \l_zchinr_article_title_alt_tl , 
  article / title ~ short  .tl_set:N    = \l_zchinr_article_title_short_tl , 
  article / title ~ short  .initial:n   = { \l_zchinr_article_title_tl } ,
  article / abstract       .tl_set:N    = \l_zchinr_article_abstract_tl , 
  article / abstract ~ alt .tl_set:N    = \l_zchinr_article_abstract_alt_tl ,
  article / category       .tl_set:N    = \l_zchinr_article_category_tl , 
  article / category       .initial:n   = { Aufsätze } , 
  article / two ~ columns  .bool_set:N  = \l_zchinr_article_twocolumn_bool , 
  article / two ~ columns  .default:n   = { true } ,
  article / two ~ columns  .initial:n   = { false } ,
  article / title ~ style  .clist_set:N = \l_zchinr_article_title_style_clist , 
  article / title ~ style  .initial:n   = { } , % category , author , abstract
  reset                    .code:n      = {
    \keys_set:nn { zchinr / setup / article } {
      author = { } ,
      author ~ short = { } ,
      author ~ copy = { } , 
      author ~ note = { } ,
      title = { } ,
      title ~ pdf = { } ,
      title ~ alt = { } ,
      title ~ short = { } ,
      abstract = { } , 
      abstract ~ alt = { } , 
      two ~ columns = false ,
      title ~ style = { }
    }
  }
}

\cs_new:Npn \__zchinr_pages_split:n #1 {
  \__zchinr_pages_split_aux:w #1 \q_stop
}

\cs_new:Npn \__zchinr_pages_split_aux:w #1 - #2 \q_stop {
  \tl_set:Nn \l_zchinr_page_first_int {#1}
  \tl_set:Nn \l_zchinr_page_last_int {#2}
}

\NewDocumentCommand { \zchinrsetup } { m } {
  \keys_set:nn { zchinr / setup } {#1}
}

\tl_new:N \l__zchinr_header_mark_even_tl
\tl_new:N \l__zchinr_header_mark_odd_tl
\tl_new:N \l__zchinr_header_mark_page_tl

\dim_new:N \l__zchinr_balance_skip_dim

\RequirePackage[english, german]{babel}

\RequirePackage{fontspec}
\setmainfont{TeX ~ Gyre ~ Pagella}
\setsansfont{TeX ~ Gyre ~ Heros ~ Cn}
\newfontfamily\zchirsansfamily[Mapping=tex-text, LetterSpace=2]{TeX ~ Gyre ~ Heros}

\RequirePackage{xeCJK} 
\XeTeXlinebreaklocale "zh" 
\XeTeXlinebreakskip = 0pt plus 1pt minus .5pt
\setCJKmainfont{Noto ~ Serif ~ CJK ~ SC}
\setCJKsansfont{Noto ~ Sans ~ CJK ~ SC}

\NewDocumentCommand{\zhs}{ m }{\textup{#1}}
\NewDocumentCommand{\zht}{ m }{\textup{#1}}

\RequirePackage[
  inner=25mm,
  outer=30mm,
  top=30mm,
  bottom=25mm,
  headsep=10mm,
  footskip=5mm
]{geometry}
\setlength{\columnsep}{7.5mm}

\cs_new:Npn \__zchinr_cropmark_nw: {
  \begin{picture}(0,0)
    \thinlines\unitlength1pt\linethickness{0.25pt}
    \put(-6,0){\line(-1,0){20}}
    \put(0,6){\line(0,1){20}}
  \end{picture}
}
\cs_new:Npn \__zchinr_cropmark_ne: {
  \begin{picture}(0,0)
    \thinlines\unitlength1pt\linethickness{0.25pt}
    \put(6,0){\line(1,0){20}}
    \put(0,6){\line(0,1){20}}
  \end{picture}
}
\cs_new:Npn \__zchinr_cropmark_sw: {
  \begin{picture}(0,0)
    \thinlines\unitlength1pt\linethickness{0.25pt}
    \put(-6,0){\line(-1,0){20}}
    \put(0,-6){\line(0,-1){20}}
  \end{picture}
}
\cs_new:Npn \__zchinr_cropmark_se: {
  \begin{picture}(0,0)
    \thinlines\unitlength1pt\linethickness{0.25pt}
    \put(6,0){\line(1,0){20}}
    \put(0,-6){\line(0,-1){20}}
  \end{picture}
}

\bool_if:NT \l__zchinr_prepress_bool {
  \geometry{
    layout=a4paper,
    paperwidth=230mm,
    paperheight=317mm,
    layouthoffset=10mm,
    layoutvoffset=10mm
  }

  \AddToHook{shipout/background}{
    \put(10mm,-10mm){\__zchinr_cropmark_nw:}
    \put(\paperwidth-10mm,-10mm){\__zchinr_cropmark_ne:}
    \put(10mm,10mm-\paperheight){\__zchinr_cropmark_sw:}
    \put(\paperwidth-10mm,10mm-\paperheight){\__zchinr_cropmark_se:}

    \special{pdf:put ~ @thispage ~ <<
      /MediaBox ~ [0.0000 ~ 0.0000 ~ 651.9690 ~ 898.5830]
      /CropBox ~ [0.0000 ~ 0.0000 ~ 651.9690 ~ 898.5830]
      /BleedBox ~ [19.8425 ~ 19.8425 ~ 632.1265 ~ 878.7405]
      /TrimBox ~ [28.3465 ~ 28.3465 ~ 623.6225 ~ 870.2365]
    >>}
  }

  \AtBeginDocument{
    \special{pdf:minorversion ~ 3}
    \special{pdf:docinfo ~ <<
      /GTS_PDFXVersion ~ (PDF/X-1:2001)
      /GTS_PDFXConformance ~ (PDF/X-1a:2001)
    >>}
    \str_if_empty:NTF { \l_zchinr_color_profile_path_tl } {
      \ClassWarning{zchinr}{No ~ color ~ profile ~ defined.}
    } {
      \special{pdf:fstream ~ @cmykdata ~ (\l_zchinr_color_profile_path_tl) ~ <<
        /N ~ 4^^J/Alternate/DeviceCMYK
      >>}
    }
    \str_if_empty:NTF { \l_zchinr_color_profile_path_tl } {
      \special{pdf:put ~ @catalog ~<<
        /PageMode ~ /UseNone
        /OutputIntents ~ [<<
          /Type ~ /OutputIntent
          /S ~ /GTS_PDFX
          /OutputConditionIdentifier ~ (Custom)
          /Info ~ (none)
          /RegistryName ~ (http://www.color.org/)
        >>]
      >>}
    } {
      \special{pdf:put ~ @catalog ~ <<
        /PageMode ~ /UseNone
        /OutputIntents ~ [<<
          /Type ~ /OutputIntent
          /S ~ /GTS_PDFX
          /DestOutputProfile ~ @cmykdata
          /OutputConditionIdentifier ~ (\l_zchinr_color_profile_id_tl)
          /Info ~ (\l_zchinr_color_profile_info_tl)
          /RegistryName ~ (http://www.color.org/)
        >>]
      >>}
    }
  }
}

\frenchspacing
\clubpenalty=10000
\widowpenalty=10000
\doublehyphendemerits=5000000

\setlength{\parindent}{15pt}
\setlength{\parskip}{2pt plus 1pt minus 1pt}

\RequirePackage{afterpage}
\RequirePackage{array}
\RequirePackage{longtable}
\RequirePackage{booktabs}
\RequirePackage{enumitem}
\RequirePackage{titletoc}
\RequirePackage[pagestyles]{titlesec}

% Style figure and table captions.
\RequirePackage[
  labelfont=bf,
  labelsep=period,
]{caption}

% Set up hyperreferences
\PassOptionsToPackage{hyphens}{url}
\bool_if:NTF \l__zchinr_prepress_bool {
  \RequirePackage[bookmarks=false]{hyperref}
  \RequirePackage{datetime}
  \AtBeginDocument{
    \NoHyper
    \hypersetup{
      pdftitle={\l_zchinr_local_zchinr_tl{} ~ \l_zchinr_issue_year_tl{}},
      pdfinfo={
        ModDate={D:\pdfdate},
        Trapped={False},      
      }
    }
  }
} {
  \RequirePackage[hidelinks]{hyperref}
}

% Style header and footer. Header stays empty while footer will contain journal information and page number. 
\newpagestyle{zchinrdefault}{
  \sethead
    [\l__zchinr_header_mark_page_tl]
    [\l__zchinr_header_mark_odd_tl]
    [\l_zchinr_local_zchinr_tl{} ~ \l_zchinr_issue_year_tl]
    {\l_zchinr_local_zchinr_tl{} ~ \l_zchinr_issue_year_tl}
    {\l__zchinr_header_mark_even_tl}
    {\l__zchinr_header_mark_page_tl}
  \setfoot{}{}{}
}
\pagestyle{zchinrdefault}

\newpagestyle{zchinrtitle}{
  \sethead{}{}{}
  \setfoot
    {\parbox{0.4\textwidth}{
      \l_zchinr_local_issn_tl \c_colon_str{} ~ \l_zchinr_issue_issn_tl \\
      \l_zchinr_local_doi_tl \c_colon_str{} ~ 
        \href{https://doi.org/\l_zchinr_article_doi_tl}{\l_zchinr_article_doi_tl}
    }}
    {\l__zchinr_header_mark_page_tl}
    {\parbox{0.4\textwidth}{\raggedleft
      \copyright{} ~ \l_zchinr_issue_year_tl{} ~ \l_zchinr_article_author_tl \\ 
      \l_zchinr_local_licence_tl \c_colon_str{} ~ 
        \href{\l_zchinr_article_licence_url_tl}{\l_zchinr_article_licence_tl}
    }}
}

\newpagestyle{zchirempty}{
  \sethead{}{}{}
  \setfoot{}{}{}
}

% Style section titles. 
\titleformat{\section}{\bfseries\fontsize{11}{13}\selectfont}{}{0pt}{}
\titleformat{\subsection}{\bfseries\fontsize{11}{13}\selectfont}{}{0pt}{}
\titleformat{\subsubsection}{\bfseries\fontsize{11}{13}\selectfont}{}{0pt}{}
\titlespacing*{\section}{0pt}{5pt plus 1pt}{5pt plus 1pt}
\titlespacing*{\subsection}{0pt}{5pt plus 1pt}{5pt plus 1pt}
\titlespacing*{\subsubsection}{0pt}{5pt plus 1pt}{5pt plus 1pt}

\titleformat{\paragraph}[runin]{\bfseries}{}{2em}{\noindent}
\titleformat{\subparagraph}[runin]{\bfseries}{}{2em}{\noindent}
\titlespacing*{\paragraph}{0pt}{5pt plus 1pt}{1em plus 5pt minus 10pt}
\titlespacing*{\subparagraph}{0pt}{5pt plus 1pt}{1em plus 5pt minus 10pt}

\titlecontents{section}[0pt]{}{\hspace*{-15pt}}{}
  {\quad\titlerule*[5pt]{.}\contentspage}
\titlecontents{subsection}[15pt]{}{\hspace*{-15pt}}{}
  {\quad\titlerule*[5pt]{.}\contentspage}
\titlecontents{subsubsection}[30pt]{}{\hspace*{-15pt}}{}
  {\quad\titlerule*[5pt]{.}\contentspage}
\titlecontents{paragraph}[45pt]{}{\hspace*{-15pt}}{}
  {\quad\titlerule*[5pt]{.}\contentspage}
\titlecontents{subparagraph}[60pt]{}{\hspace*{-15pt}}{}
  {\quad\titlerule*[5pt]{.}\contentspage}

% Style footnotes.
\RenewDocumentCommand { \@makefntext } { +m } {
  \parbox[t]{2em}{\@thefnmark}
  \parbox[t]{\dimexpr\columnwidth-2em}{#1}
}
\RenewDocumentCommand { \footnoterule } { } { }
\setlength{\skip\footins}{15pt plus 5pt minus 5pt}

% Create listing environment for references. 
\newlist{references}{enumerate}{1}
\setlist[references]{label*=[\arabic*]}

% Create command for automatically linked e-mail addresses. 
\NewDocumentCommand { \email } { m } {
  \href{mailto:#1}{\protect\nolinkurl{#1}}
}

% Create documentation environment. 
\NewDocumentEnvironment { documentation } { } {
  \renewcommand{\arraystretch}{1.25}
  \noindent\begin{longtable}{
    @{} >{\hskip 2em} p{54mm}
    @{\hskip 7mm} >{\hskip 15pt} p{\dimexpr\textwidth - 54mm - 7mm}
    @{}
  }
} {
  \end{longtable}
}

\NewDocumentCommand { \tablehead } { O{c} O{} m o o } {
  \IfNoValueF{#4}{
    \phantomsection
    \addcontentsline{toc}{#4}{\IfNoValueTF{#5}{#3}{#5}}
  }
  \vspace{-\baselineskip}
  \begin{center}
    \str_if_eq:nnT {#1} { l } { \raggedright }
    \str_if_eq:nnT {#1} { r } { \raggedleft }
    \str_if_eq:nnF {#2} { n } { \bfseries } #3
  \end{center}
}

% Create addresses environment.
\NewDocumentEnvironment { addresses } { } {
  \setlength{\arrayrulewidth}{1pt}
  \renewcommand{\arraystretch}{1.75}
  \noindent\begin{longtable}{
    @{} p{80.75mm}
    @{\hskip 7mm} p{\dimexpr\textwidth - 88.75mm - 7mm}
    @{}
  }
} {
  \end{longtable}
}

% Create imprint environment.
\NewDocumentEnvironment { imprint } { } {
  \renewcommand{\arraystretch}{1.25}
  \noindent\begin{longtable}{
    @{} p{30mm}
    @{\hskip 7mm} p{\dimexpr\textwidth - 30mm - 7mm - 4mm - 35mm}
    @{\hskip 4mm} >{\raggedleft\arraybackslash} p{35mm}
    @{}
  }
} {
  \end{longtable}
}

% Create subcontents environment.
\NewDocumentEnvironment { subcontents } { } {
  \list{}{\leftmargin=40mm \rightmargin=25mm}
  \item\relax
} {
  \endlist
}

% Create commands for TOC. 
\NewDocumentCommand { \tocsection } { m }{\bigskip\noindent{\Large\scshape\bfseries #1}}

\newlist{toc}{itemize}{1}
\setlist[toc]{leftmargin=15mm, label*={}}

% Redefine abstract environment.
\RenewDocumentEnvironment { abstract } { O{\textwidth} } {
  \begin{minipage}[t]{#1}
  \setlength{\parindent}{15pt}
  \noindent
} {
  \end{minipage}
}

\NewDocumentCommand { \zchinrprinttitle } { } {
  \clearpage
  \thispagestyle{zchinrtitle}
  \setcounter{footnote}{0}
  \startcontents
  
  \bool_if:NTF \l_zchinr_pagination_bool {
    \tl_set:Nn \l__zchinr_header_mark_page_tl { \thepage }
  } {
    \tl_set:Nn \l__zchinr_header_mark_page_tl { }
  }
  
  \tl_if_empty:NTF \l_zchinr_article_title_short_tl {
    \tl_set_eq:NN \l__zchinr_header_mark_even_tl \l_zchinr_article_category_tl
    \tl_set_eq:NN \l__zchinr_header_mark_odd_tl \l_zchinr_article_category_tl  
  } {
    \tl_if_empty:NTF \l_zchinr_article_author_short_tl {
      \tl_set_eq:NN \l__zchinr_header_mark_even_tl \l_zchinr_article_title_short_tl
      \tl_set_eq:NN \l__zchinr_header_mark_odd_tl \l_zchinr_article_title_short_tl
    } {
      \tl_set:Nn \l__zchinr_header_mark_even_tl {
        { \itshape \l_zchinr_article_author_short_tl , } ~ \l_zchinr_article_title_short_tl
      }
      \tl_set:Nn \l__zchinr_header_mark_odd_tl {
        { \itshape \l_zchinr_article_author_short_tl , } ~ \l_zchinr_article_title_short_tl
      }
    }
  }

  \renewcommand*{\thefootnote}{\ensuremath{*}}
  \bool_if:NTF \l_zchinr_article_twocolumn_bool {
    \twocolumn[
      \__zchir_article_title:
    ]
    \sloppy\flushbottom
    \enlargethispage{-10mm}
    \afterpage{\enlargethispage{-10mm}}
  } {
    \onecolumn
    \__zchir_article_title:
    \enlargethispage{-10mm}
  }
  \tl_if_empty:NF \l_zchinr_article_author_note_tl {
    \footnotetext{\l_zchinr_article_author_note_tl}
  }
  \renewcommand*{\thefootnote}{\arabic{footnote}}
  \@afterindentfalse\@afterheading
}

\cs_new:Npn \__zchir_article_author: {
  \l_zchinr_article_author_tl
  \tl_if_empty:NF \l_zchinr_article_author_note_tl {
    \, \footnotemark[0]
  }
}

\cs_new:Npn \__zchir_article_title: {
  \clist_if_in:NnT \l_zchinr_article_title_style_clist { category } {
    \noindent\rule{\textwidth}{1pt}
    \par\vskip 2mm
    \int_if_even:nTF { \l_zchinr_article_page_first_int } {
      \noindent\parbox[t]{\textwidth}{
        \raggedleft\huge\scshape\bfseries\l_zchinr_article_category_tl 
      }\par
    } {
      \noindent\parbox[t]{\textwidth}{
        \raggedright\huge\scshape\bfseries\l_zchinr_article_category_tl
      }\par
    }
    \noindent\rule{\textwidth}{1pt}
    \par\vskip 7mm
  }

  \bool_if:NF \l__zchinr_standalone_bool {
    \phantomsection
    \addcontentsline{toc}{chapter}{\l_zchinr_article_title_tl}
  }

  \noindent\parbox[t][\l_zchinr_title_distance_dim][t]{\textwidth}{
    \raggedright\huge\l_zchinr_article_title_tl
    \clist_if_in:NnT \l_zchinr_article_title_style_clist { author } {
      \par\bigskip
      \fontsize{11}{13}\selectfont\itshape \__zchir_article_author:
    }
  }
  \par\vskip 4mm\bigskip

  \clist_if_in:NnT \l_zchinr_article_title_style_clist { abstract } {
    \noindent{\itshape\bfseries\fontsize{11}{13}\selectfont\l_zchinr_local_abstract_tl }
    \par\vskip 8pt
    \str_if_empty:NTF \l_zchinr_article_abstract_alt_tl {
      \noindent\begin{abstract}
        \itshape\l_zchinr_article_abstract_tl
      \end{abstract}
    } {
      \noindent\begin{abstract}[\columnwidth]
        \itshape\l_zchinr_article_abstract_tl
      \end{abstract}
      \hskip\columnsep
      \begin{abstract}[\columnwidth]
        \itshape{\bfseries\l_zchinr_article_title_alt_tl} ~ --- ~
        \l_zchinr_article_abstract_alt_tl
      \end{abstract}
    } 
    \par\vskip 12mm
  }

  \clist_if_in:NnT \l_zchinr_article_title_style_clist { contents } {
    \begin{subcontents}
      \printcontents{l}{1}{}
    \end{subcontents}
    \par\vskip 12mm
  }
}

% EOF