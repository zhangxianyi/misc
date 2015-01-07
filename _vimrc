set nocompatible
source $VIMRUNTIME/vimrc_example.vim
source $VIMRUNTIME/mswin.vim
behave mswin
set diffexpr=MyDiff()
function! MyDiff()
  let opt = '-a --binary '
  if (&ic == 1)
    let opt = opt . '-i '
  endif
  if &diffopt =~ 'icase' | let opt = opt . '-i ' | endif
  if &diffopt =~ 'iwhite' | let opt = opt . '-b ' | endif
  let arg1 = v:fname_in
  if arg1 =~ ' ' | let arg1 = '"' . arg1 . '"' | endif
  let arg2 = v:fname_new
  if arg2 =~ ' ' | let arg2 = '"' . arg2 . '"' | endif
  let arg3 = v:fname_out
  if arg3 =~ ' ' | let arg3 = '"' . arg3 . '"' | endif
  if &sh =~ '\<cmd'
    silent execute '!""D:\skill\Apps\Vim\vim73\diff.exe" ' . opt . arg1 . ' ' . arg2 . ' > ' . arg3 . '"'
  else
    silent execute '!$VIM\vim73\diff" ' . opt . arg1 . ' ' . arg2 . ' > ' . arg3
  endif
  execute("syntax off")
endfunction

set diffopt+=icase
set diffopt+=iwhite
set encoding=cp936
" set encoding=utf-8
set termencoding=cp936
set fileencodings=ucs-bom,utf-8,cp936
" set fileencodings=cp936,ucs-bom,utf-8
" set fileencodings=cp936,ucs-bom,utf-8

if filereadable($VIM.'/map.vim')
  source $VIM/map.vim
endif

" set shellslash
" set noshellslash
set visualbell

" consult vim help setting 78
" set tw=78
set lbr

set noequalalways
set nobackup
set nowritebackup
set nu noscs nowrap
set ic
" set noic
set bsdir=buffer
set wak=no
set viminfo='20,/10,<0,s0,:0
set shiftwidth=4
set cindent
set matchpairs+=<:>
" set guifont=Envy_Code_R:h11:cANSI
" set guifont=Consolas:h10:cANSI
" set guifont=Droid_Sans_Mono:h10:cANSI
set guifont=Consolas:h11:cANSI
"set guifont=Arial_monospaced_for_SAP:h11:cANSI
"set guifont=Bitstream_Vera_Sans_Mono:h10:cANSI
"set gfw=幼圆:h10:cGB2312
" set guifont=Droid_Sans_Mono:h13:cANSI

" set guioptions=aem
set guioptions=Ae
set winminheight=0
set cmdwinheight=3
set cmdheight=2
set winheight=1

set isfname-==
set isfname+=(,)
set iskeyword+=_

set grepprg=findstr\ /n\ /s\ /I\ 

set helplang=en,cn

hi StatusLine	gui=bold	guifg=yellow	guibg=black
hi Foo_MyTagListTagName	gui=bold	guifg=black	guibg=green
hi MyTagListTagName	gui=bold	guifg=black	guibg=green
hi Foo_CurLine_Highlight gui=bold	guifg=black	guibg=yellow

" show whitespace at the end of a line, highlight must set before match
highlight WhitespaceEOL ctermbg=blue guibg=gray
match WhitespaceEOL /\s\+$/


" let Tlist_Use_Right_Window=1
" let Tlist_WinHeight = 10
" let Tlist_Use_Horiz_Window=0
" let Tlist_File_Fold_Auto_Close=1
" let g:Tlist_WinWidth=37

if !exists("g:FS")
  if has("win32") && &ssl==0
    let g:FS = '\'
  else
    let g:FS = '/'
  endif
endif

augroup Binary
  au!
  au BufReadPre  *.xls,*.jpg,*.a,*.dat,*.bin,*.o,*.exe setlocal bin
  au BufReadPost *.xls,*.jpg,*.a,*.dat,*.bin,*.o,*.exe if &bin | %!xxd
  au BufReadPost *.xls,*.jpg,*.a,*.dat,*.bin,*.o,*.exe set ft=xxd | endif
  au BufWritePre *.xls,*.jpg,*.a,*.dat,*.bin,*.o,*.exe if &bin | %!xxd -r
  au BufWritePre *.xls,*.jpg,*.a,*.dat,*.bin,*.o,*.exe endif
  au BufWritePost *.xls,*.jpg,*.a,*.dat,*.bin,*.o,*.exe if &bin | %!xxd
  au BufWritePost *.xls,*.jpg,*.a,*.dat,*.bin,*.o,*.exe set nomod | endif
augroup END

augroup code
  au BufNewFile,BufReadPost,BufWritePost		*.c,*.h		setlocal fenc=cp936
  au BufNewFile,BufReadPost,BufWritePost		*.c,*.h		setlocal ff=unix
augroup END

" augroup cm
"   au BufNewFile,BufRead         *.cm            set ft=cpp
" augroup END

augroup reg
  au!
  au BufReadPre                   *.reg  set enc=utf-8
  au BufReadPre                   *.reg  set fileencodings=ucs-bom,utf-8,cp936
  au BufNewFile,BufReadPost,BufWritePost		*.reg	set enc=cp936
  au BufNewFile,BufReadPost,BufWritePost		*.reg	set fenc=cp936
augroup END " augroup reg

" autocmd VimEnter * nested if
"       \ filereadable($VIM."/Session.vim") |
"       \ source $VIM/Session.vim |elseif
  "       \ argc() == 0 &&
  "       \ bufname("%") == "" |
  "       \   exe "normal! `0" |
  "       \ endif

" vim关闭时执行保存当前vim的窗口的信息到session文件中.
" 保存viminfo信息到文件$HOME/_viminfo
" autocmd VimLeave * let save=confirm("Save session?", "&Yes\n&No\n&Delete", 1) |
"       \ if save == 1 | mksession! $HOME/Session.vim |
"       \ elseif save == 3 | silent execute("!start rm -f " .$HOME.'/Session.vim') |
"       \ endif

" 启动时如果没有指定参数则恢复之前打开的session和_viminfo
autocmd VimEnter * nested if
      \ argc() == 0 &&
      \ bufname("%") == "" |
      \ exe "normal! `0" |
      \ endif

" 当读入文件到buffer时, 将光标定位到此文件最后光标所在的位置.
" 现在发现`可以移动到许多地方. 而g`就是jump的命令.
autocmd BufReadPost *
      \ if line("'\"") > 1 && line("'\"") <= line("$") |
      \   exe "normal! g`\"" |
      \ endif

if !exists("g:UseWinCopy")
  let g:UseWinCopy = 0
endif

if !exists("g:Execute_Extension")
  " let g:Execute_Extension=',dll,PR,ico,html,7z,htm,hlp,tar,ps,pdf,exe,ttf,gif,dvi,torrent,doc,jpg,JPG,bmp,rar,png,chm,zip,ppt,avi,rm,rmvb,hlp,mp3,wav,'
  let g:Execute_Extension=',dll,PR,ico,7z,hlp,tar,ps,pdf,exe,ttf,gif,dvi,torrent,doc,jpg,JPG,bmp,rar,png,chm,zip,ppt,avi,rm,rmvb,hlp,mp3,wav,'
endif

if !exists("g:TagSep ")
  let g:TagSep  = ','
endif

if !exists("g:needMakePrompt")
  let g:needMakePrompt = 1
endif

set nosecure

set fo+=m

set ssop+=resize
set ssop-=curdir

" if !exists('g:Tlist_WinScale')
"   let g:Tlist_WinScale = 0.2
" endif

" use for F:\PortableApps\vim\\vimfiles\ftplugin\java.vim
if !exists('g:WinFS')
  let g:WinFS = '\'
endif

" used in F:\skill\vim\ftplugin\java\SetTag\Java_HasImport.vim
if !exists("g:Java_Default_Tags")
  let g:Java_Default_Tags = ''
endif

let g:fe_es_exe = 'D:\skill\Apps\Everything\es.exe'
set mouse=n
"let g:pydoc_cmd = 'python C:\Python27\Lib\pydoc.py'
set spr
let &shm="a"
let no_buffers_menu = 1
" set mousemodel=extend
" set mousemodel&vim
set mouse=a

"au BufRead,BufNewFile _pentadactylrc set filetype=pentadactyl
" disable explore.vim currently
let g:loaded_netrw = 1
let g:loaded_netrwPlugin = 1
let g:loaded_netrwFileHandlers = 1
" Nread scp://afj@192.168.1.105//usr/local/apache2/htdocs
" Nread scp://root@192.168.1.105//
" call NetUserPass("userid","passwd")
" call NetUserPass("root@192.168.1.105","afjafj")
" let g:netrw_cygwin = 0
" let g:netrw_scp_cmd='D:\skill\Apps\putty\putty6.0_nd1.14\pscp.exe -pw afjafj '
" let loaded_explorer=1
autocmd BufEnter * lcd %:p:h

if !exists("g:mobiledisk_label")
  " let g:mobiledisk_label=strpart($VIM, 0, 1)
  let g:mobiledisk_label='e'
endif

runtime bundle/vim-pathogen/autoload/pathogen.vim
call pathogen#infect()

" PmenuSel       xxx ctermbg=7 guibg=Grey
" 设置pop menu selection的颜色
highlight PmenuSel guifg=black guibg=red
let indexer_indexerListFilename='D:/skill/Apps/Vim/vimfiles/bundle/indexer/.indexer_files'
let indexer_tagsDirname='D:/skill/Apps/Vim/vimfiles/bundle/indexer/.indexer_files_tags/'

" for cSyntaxAfter
" autocmd! BufRead,BufNewFile,BufEnter *.{c,cpp,h,javascript} call CSyntaxAfter()

" let g:indexer_debugLogFilename='D:/PortableApps/vim/vimfiles/bundle/indexer/debug.out'
" let g:indexer_debugLogFilename=""
" let g:indexer_debugLogLevel=3
" fn.name
" let g:EasyMotion_mapping_w = '_w'
" let g:EasyMotion_leader_key = '<Leader>'
" let g:EasyMotion_leader_key = '\'
nnoremap <M-,> <C-o>
nnoremap <M-.> <C-i>
unmap <C-a>
" let g:pydiction_location = 'D:\PortableApps\vim\vimfiles\bundle\pydiction-1.2.1\complete-dict'
let g:emdiction_location = 'D:\PortableApps\vim\vimfiles\bundle\emdiction\em.dict'
set ve=all
au BufReadCmd   *.epub      call zip#Browse(expand("<amatch>"))
let g:siprg = 'D:\skill\Apps\insight3\Insight3.exe'

filetype off
"call vundle#rc('D:/PortableApps/vim/vimfiles/bundle')

" let Vundle manage Vundle
" required!
"Bundle 'gmarik/vundle'
"Bundle 'scrooloose/syntastic'
" script-names.vim-scripts.org.json not support yet
" Bundle 'mbbill/undotree'
"Bundle 'jelera/vim-javascript-syntax'
"Bundle 'vimprj'
"colorscheme github

filetype plugin indent on
" set fdm=indent
let g:AutoClosePairs = {"<":">", '(': ')', '{': '}', '[': ']', '"': '"', "'": "'"}
let g:NERDTreeWinPos = "right"
" highlight DiffChange cterm=none ctermfg=black ctermbg=LightGreen gui=none guifg=bg guibg=LightGreen
" highlight DiffText cterm=none ctermfg=black ctermbg=Red gui=none guifg=bg guibg=Red
" highlight DiffAdd ctermbg=22 gui=NONE guifg=bg guibg=Green
" highlight DiffDelete ctermbg=52 gui=NONE guifg=bg guibg=Red
" highlight DiffChange ctermbg=53 gui=NONE guifg=bg guibg=Yellow
" highlight DiffText ctermbg=57 gui=NONE guifg=bg guibg=Magenta
highlight DiffText guibg=Orange
"highlight Constant term=underline ctermfg=4 guifg=Magenta
highlight Constant guifg=red
"let g:mwDefaultHighlightingPalette = 'extended'
let g:mwDefaultHighlightingPalette = 'maximum'
let g:NERDTreeChDirMode = 2
let NERDTreeIgnore=['\.$', '\.\.$', '\~$', '.swp$']
let NERDTreeShowHidden=1
au BufNewFile,BufRead *.handlebars set filetype=html
