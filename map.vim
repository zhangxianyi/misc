nnoremap <M-space> :simalt ~<CR>
"nnoremap vsp :vsp<cr>

"cnoremap <M-d> <C-d>
cnoremap <M-d> <Del>

cnoremap <M-;> ;
cnoremap <M-f> <cr>
" cnoremap <M-w> <C-W>
cnoremap <M-w> <C-w>
cnoremap <M-,> <C-Left>
cnoremap <M-.> <C-Right>
cnoremap <M-H> <C-Left>
cnoremap <M-L> <C-Right>
cnoremap <M-b> <C-left>
cnoremap <M-c> <C-u>
cnoremap <M-u> <C-u>
cnoremap <M-l> <Right>
cnoremap <M-h> <Left>
cnoremap <M-j> <Down>
cnoremap <M-k> <Up>
cnoremap <M-a> <End>
cnoremap <M-i> <Home>
cnoremap <M-r> <C-r>r
cnoremap <M-v> <C-R>*
cnoremap <M-m> <cr>
cnoremap <M-x> <del>
cnoremap ; <cr>

map vfe :FSplit<cr>/
map vfs :FSplit<cr>/
imap <M--> <Esc><C-w>_a
map <M--> <C-w>_
map <M-0> <C-W>o
nnoremap <silent> <M-s> :silent w<cr>
" map <M-q> :q<cr>
map <M-q> :if (&mod==0 \|\| expand("%:p")=="") \|q!\|endif<cr>

map <M-n> :new<cr>
" map <M->> <C-w>>
" map <M-<> <C-w><
" map <M-,> <C-w>-
" map <M-.> <C-w>+
map <M->> <C-w>+
map <M-<> <C-w>-
map <M-,> :tabp<cr>
map <M-.> :tabn<cr>
map <M-w> <C-w>

map gl "*y$
map gr "*yG
" tab nag
" map gt :let @*=expand("%:t")<CR>
" map ge :let @*=expand("%:t:r")<CR>
map ge :let @*=expand("%:t")<CR>
map <M-O> O
map \q :let @*=expand("<cword>")<cr>
map q "*yw
map <M-M> :BufExplorer<CR>
map t $
map <M-]> :silent! execute('lcd ' . escape(GetDirectoryName(),  '%#!*?'))<cr>:silent! execute "normal \<C-]>"<cr>
nnoremap <C-[> <C-t>
map <Tab> :let @*=getline(".")<CR>
map ej :<C-v><CR>
map <M-a> A
map <M-v> <C-U>
nnoremap <M-f> <C-F>
map <M-g> G
cnoremap <M-g> <C-R>=substitute(getline("."),'\/','','')<cr>
map <M-x> :simalt ~n<cr>
" nnoremap <M-3> g#
nnoremap <M-9> :AS<cr>
nnoremap <M-=> <C-W>=
nnoremap <F1> K
nnoremap vv :<C-F>
" nnoremap <M-z> zz
nnoremap <Space> :set ic<cr>/
nnoremap ; :
nnoremap vso :!start "c:\Program Files\OpenOffice\OpenOffice.org 3\program\soffice.exe"<cr>
nnoremap vsd :!start C:\Apps\Netease\网易闪电邮\start.exe<cr>
nnoremap vbc :!start "d:\PortableApps\Beyond Compare\Beyond Compare 3\BCompare.exe"<cr>
nnoremap vim :silent !start d:\skill\apps\vim\vim74\gvim.exe<cr>
" nnoremap vpu :silent !start H:\skill\Term\putty\release\0.60\PUTTY.EXE -l afj -load "debian-hostonly"<cr>
" nnoremap vpn :silent !start d:\PortableApps\putty\0.60\cn\putty.exe -l afj -pw afj -load "f13-192.168.1.105"<cr>
" nnoremap vpn :silent !start d:\PortableApps\putty\0.60\cn\putty.exe -l afj -pw afj -load "f13-1.1.1.3"<cr>
nnoremap vpn :silent !start d:\PortableApps\putty\0.60\cn\putty.exe -l afj -pw afj -load "f13-10.0.0.3"<cr>
" nnoremap vpn :silent !start d:\PortableApps\putty\0.60\cn\putty.exe -l afj -pw afj -load "f13-192.168.205.128"<cr>
" d:\PortableApps\putty\0.60\cn\putty.exe -l root -pw afjafj -load "f13-1"
" C:\Program Files\WinHTTrack\WinHTTrack.exe"
" nnoremap vpn :silent !start d:\PortableApps\putty\0.60\cn\putty.exe -l root -pw afj -load "debian505"<cr>
" nnoremap vu :silent !start d:\PortableApps\putty\0.60\cn\putty.exe -l root -pw afj -load "debian505"<cr>

inoremap <M-t> <Esc>:let @*=expand("%:t:r")<CR>a
" inoremap <M-h> <Esc>:let @*=expand("%:h")<CR>a
imap <M-q> <Esc>:q<cr>
inoremap <M-u> <C-u>
inoremap <M-8> *
imap <M-]> <Esc>}i
imap <M-g> <C-End>

imap <M-x> <Esc>xa
imap <M-v> <C-v>
imap <M-p> <C-v>
cmap <M-8> *
cmap <M-4> $
cmap <M-6> ^
vmap <M-4> $
imap <M-4> $
cmap <M-p> <C-v>
let g:zxycomment="---------------------------------------------------------"
inoremap <M-d> [Date]: <C-R>=strftime("%x")
inoremap <M-c> <Esc>S<Home>---------------------------------------------------------
inoremap <M-k> <Up>
inoremap <M-j> <Down>
inoremap <M-h> <Left>
inoremap <M-l> <Right>
inoremap <M-i> <Home>
inoremap <M-a> <End>
inoremap <M-H> <C-Left>
inoremap <M-L> <C-Right>
inoremap <M-r> <C-R>r
inoremap <M-n> <Esc>:new<cr>
inoremap <M-o> <Esc>o
inoremap <M-O> <Esc>O
inoremap <M-w> <C-W>
inoremap <M-s> <C-U>
inoremap <M-b> <C-left>

nnoremap <rightmouse> <Nop>
inoremap <MiddleMouse> <Nop>
inoremap <rightmouse> <Nop>

vnoremap <Space> /

" inoremap <M-?> <Esc>:winpos 444 440<cr>a
" inoremap <M-Q> <Esc>:winpos -1 -1<cr>a
" inoremap <M-q> <Esc>:winpos -1 -1<cr>a
" inoremap <M-P> <Esc>:winpos 176 -1<cr>a
" inoremap <M-p> <Esc>:winpos 176 -1<cr>a
" inoremap <M-z> <Esc>:winpos 176 440<cr>a

" nnoremap <M-?> :winpos 444 440<cr>
" nnoremap <M-/> :winpos 444 440<cr>
" nnoremap <M-Q> :winpos -1 -1<cr>
" nnoremap <M-p> :winpos 176 -1<cr>
" nnoremap <M-'> <C-]>
" nnoremap <M-Z> :winpos 176 440<cr>
" nnoremap <M-z> :winsize 90 18<cr>:winpos -4 410<cr>
" nnoremap <M-/> :winsize 90 18<cr>:winpos 620 410<cr>
" nnoremap <M-?> :winsize 90 18<cr>:winpos 320 410<cr>
" nnoremap <M-p> :winsize 90 18<cr>:winpos 620 -4<cr>
" nnoremap <M-z> :winpos 176 440<cr>

cmap <M-c> <C-c>

hi Foo_CurLine_Highlight gui=bold	guifg=black	guibg=yellow

nnoremap vkmp :silent !start c:\Program Files\The KMPlayer\KMPlayer.exe<cr>
" nnoremap vmpc :silent !start F:\skill\media\Video Players\win32\mplayerc\release\MPC-HC\1.3.1249.0\MPC-Homecinema.1.3.1249.0.(x86)\mpc-hc.exe<cr>
nnoremap vmpc :silent !start c:\Program Files\K-Lite Codec Pack\Media Player Classic\mpc-hc.exe<cr>
" nnoremap vmpc :silent !start F:\skill\media\Video Players\win32\mplayerc\release\MPC-HC\1.4.2499\<cr>
" nnoremap vad :silent !start C:\Program Files\Adobe\Reader 9.0\Reader\AcroRd32.exe<cr>
" nnoremap vad :silent !start C:\Program Files\Adobe\Acrobat 7.0\Acrobat\Acrobat.exe<cr>
nnoremap vad :silent !start C:\Program Files\Adobe\Reader 10.0\Reader\AcroRd32.exe<cr>
" nnoremap vfb :silent !start f:\PortableApps\foobar2000\1.0.3\foobar2000.exe<cr>
nnoremap vfb :silent !start D:\skill\Apps\foobar2000\foobar2000.exe<cr>
" nnoremap vff :silent !start d:\PortableApps\FirefoxPortable\v3.6.8\FirefoxPortable.exe<cr>
nnoremap vff :silent !start C:\Program Files\Mozilla Firefox\firefox.exe<cr>
" nnoremap vsg :silent !start h:\skill\network\anonymity.network\google\apps\hyk-proxy\release\0.9.0\hyk-proxy-client-0.9.0\bin\startgui.bat<cr>
" nnoremap vsg :silent !start D:\skill\network\proxy\hyk-proxy-client-0.9.0\bin\startgui.bat<cr>
" nnoremap vsg :silent !start d:\skill\network\proxy\Google.App.Engine\hyk-proxy\client\java\0.9.4.1\hyk-proxy-0.9.4.1\bin\start.bat<cr>
nnoremap vsg :silent !start D:\skill\network\proxy\Google.App.Engine\snova\gsnova-0.22.1\gsnova.exe<cr>
" nnoremap vsg :silent !start d:\skill\network\proxy\Google.App.Engine\hyk-proxy\client\java\0.9.1\hyk-proxy-client-0.9.1\bin\start.bat<cr>
" C:\Program Files\Microsoft Office\Visio11\VISIO.EXE
nnoremap vie :silent !start c:/Program Files/Internet Explorer/iexplore.exe<cr>
nnoremap vfx :silent !start C:/Program Files/China Mobile/Fetion/Fetion.exe<cr>
nnoremap vtm :silent !start rundll32.exe shell32.dll,Control_RunDLLAsUser "timedate.cpl"<cr>
nnoremap vma :silent !start rundll32.exe shell32.dll,Control_RunDLLAsUser "main.cpl",,3<cr>
nnoremap vss :silent !start rundll32.exe shell32.dll,Control_RunDLLAsUser "sysdm.cpl",,3<cr>
nnoremap vca :silent !start D:/skill/Apps/Calibre2/calibre.exe<cr>
nnoremap vce :silent !start D:\skill\Apps\Calibre2\ebook-viewer.exe<cr>
nnoremap vnc :silent !start rundll32.exe shell32.dll,Control_RunDLLAsUser "ncpa.cpl"<cr>
nnoremap vin :silent !start rundll32.exe shell32.dll,Control_RunDLLAsUser "intl.cpl",,1<cr>
nnoremap vse :silent !start C:/WINDOWS/system32/mmc.exe "c:/windows/system32/services.msc"<cr>
nnoremap vas :silent !start c:\Apps\AutoHotkey\1.0.48.05.L61\AU3_Spy.exe<cr>
nnoremap vmm :silent !start rundll32.exe shell32.dll,Control_RunDLLAsUser "mmsys.cpl",,0<cr>
nnoremap vxd :silent !start D:\skill\Apps\Kingsoft\PowerWordDict\XDict.exe<cr>
nnoremap vqq :silent !start D:\skill\Apps\Tencent\Bin\QQ.exe<cr>
" nnoremap vqq :silent !start C:/Program Files/Tencent/QQ/bin/QQ.exe<cr>
nnoremap val :silent !start c:\Program Files\AliWangWang\AliIM.exe<cr>
nnoremap vao :silent !start c:\Program Files\Ashampoo\Ashampoo Burning Studio 6 FREE\burningstudio.exe<cr>
" nnoremap vao :silent !start D:\skill\Apps\Calibre\calibre.exe<cr>
nnoremap vxl :silent !start c:\Program Files\Thunder Network\MiniThunder\MiniThunder.exe<cr>
" inoremap <M-G> <C-r>=substitute(@/, '\\<\\|\\>', '', 'g')<cr>
" cnoremap <M-G> <C-r>=substitute(@/, '\\<\\|\\>', '', 'g')<cr>
" inoremap <M-/> <C-r>/
" cmap <M-/> <C-r>/
inoremap <M-/> <C-r>=substitute(@/, '\\<\\|\\>', '', 'g')<cr>
cnoremap <M-/> <C-r>=substitute(@/, '\\<\\|\\>', '', 'g')<cr>

" nnoremap <MiddleMouse> <Nop>
" nnoremap <MiddleMouse> <LeftMouse>:call HighlightCurWord()<cr>
" nnoremap <M-8> :call HighlightCurWord()<cr>
" nnoremap <M-8> :set noic<cr>:let @/='\<' . expand("<cword>") . '\>'<cr>
" nnoremap <M-8> *

" func! HighlightCurWord()
"   let @/='\<' . expand("<cword>") . '\>'
"   set noic
" endf

nnoremap <2-MiddleMouse> <NOP>
" nnoremap <MiddleMouse> <LeftMouse>*N
nnoremap vqd :silent !start c:\Program Files\Tencent\QQDownload\QQDownload.exe<cr>
nnoremap vor :silent !start c:\Program Files\Orbitdownloader\orbitdm.exe<cr>
nnoremap vgr :silent !start C:\Program Files\Orbitdownloader\Grab.exe<cr>
nnoremap vta :silent !start c:/windows/system32/taskmgr.exe<cr>
nnoremap vts :silent !start c:/windows/system32/taskswitch.exe<cr>
nnoremap vms :silent !start mspaint.exe<cr>
"nnoremap vms :silent !start D:\skill\Apps\IrfanView\i_view32.exe<cr>
" nnoremap vpe :silent !start F:\skill\utilities\win32\sysinternals\process explorer\release\v12.04\ProcessExplorer\procexp.exe<cr>
nnoremap <silent> ' `

nmap ,, !!boxes -d commonf
vmap ,, !boxes -d commonf
nmap ,f !!boxes -d commonf
vmap ,f !boxes -d commonf
nmap ,b !!boxes -d commonf
vmap ,b !boxes -d commonf

au BufRead boxes.cfg set filetype=boxes

imap <M-3> #
map ,p :echo system("netstat -an\|grep 166.111.48.32")<CR>

cmap <M-3> #
cmap <M-2> @
cmap <M-1> !
imap <M-1> !
cmap <M-5> %
vmap <M-5> %
imap <M-5> %
imap <M-6> ^
vmap <M-6> ^
nmap ,r !!boxes -d common -r<cr>
vmap ,r !boxes -d common -r<cr>
cmap <A--> _
imap <A--> _
cmap <A-7> &
cmap <A-=> +
inoremap <M-2> @

" nnoremap <2-LeftMouse> zO
nnoremap vmd :silent !start C:\Windows\hh.exe "C:\Program Files\Microsoft Visual Studio\MSDN\2001OCT\1033\MSDN130.COL"<cr>
nnoremap vmv :silent !start C:\Program Files\Microsoft Visual Studio\COMMON\MSDev98\Bin\MSDEV.EXE<cr>
nnoremap vll :silent !start C:\Program Files\Lingoes\Translator2\Lingoes.exe<cr>
nnoremap vso :silent !start c:\Program Files\OpenOffice\OpenOffice.org 3\program\soffice.exe<cr>
" nnoremap vcr :silent !start h:\PortableApps\Cryptload\1.1.8\CryptLoad.exe<cr>
nnoremap vcr :silent !start d:\PortableApps\Cryptload\1.1.8\CryptLoad.exe<cr>
" nnoremap vsg :silent !start F:\skill\network\anonymity.network\google\apps\hyk-proxy\release\0.8.6\hyk-proxy-client-0.8.6\bin\startgui.bat<cr>
" nnoremap vsg :silent !start F:\skill\network\anonymity.network\google\apps\hyk-proxy\release\0.9.0\hyk-proxy-client-0.9.0\bin\startgui.bat<cr>

" autocmd GUIEnter * winpos 176 248|winsize 80 15
" 从kmp的窗口大小的快捷键得到启示.
" map <M-!> :winsize 80 15<cr>:winpos 176 248<cr>
" map <C-M-1> :winsize 80 15<cr>:winpos 176 248<cr>
" winwidth(0)
" echo winwidth(0)
" winheight(0)
" echo winheight(0)
" map <M-@> :winsize 118 22<cr>:winpos 176 248<cr>
" map <M-`> :winsize 118 22<cr>:winpos 176 248<cr>
" map <M-~> :winsize 90 15<cr>:winpos 176 248<cr>
" map <C-S-h> :winsize 90 95<cr>:winpos -4 -4<cr>
" map <M-!> :winsize 90 18<cr>:winpos -4 380<cr>
" map <C-S-l> :winsize 90 95<cr>:winpos 620 -4<cr>
" map <M-C-2> :winsize 118 22<cr>:winpos 176 248<cr>
" 如果已经在运行，则提示已经被使用
" noremap vmw :!start "c:\Program Files\VMware\VMware Player\vmplayer.exe" "d:\skill\Virtual.Machine\Machines\debian\vmware\debian505\debian505.vmx"<cr>
noremap vmw :!start "c:\Program Files\VMware\VMware Player\vmplayer.exe"<cr>
" nnoremap <silent> <F9> :echohl WarningMsg<cr>:echo system('df -h')<cr>:echohl None
nnoremap <silent> <F9> :%s/\s\+$//<cr>
"nnoremap <silent> <F10> :diffthis<cr>
nnoremap <silent> <F10> :%s/[‘’“”]/'/g<cr>
" nnoremap K :!start mh <cword><cr>
" nnoremap <MiddleMouse> <LeftMouse>:let @/='\<' . expand("<cword>") . '\>'<cr>
" nnoremap <MiddleMouse> <LeftMouse>:let @/=expand("<cword>")<cr>
" nnoremap <MiddleMouse> <Leader>m
map <M-1> <C-^>
"map <M-2> <C-^>
map <M-3> <C-^>
" nnoremap <M-3> :set noic<CR>:let @/='\<' . expand("<cword>") . '\>'<CR>
" nnoremap <M-4> <C-w>o
noremap <M-4> $
nmap <M-5> %
" nnoremap <M-8> :set ic<cr>:let @/=expand("<cword>")<cr>
" nnoremap <M-3> :set ic<cr>:let @/='\<'.expand("<cword>").'\>'<cr>
"nnoremap <M-8> :set ic<cr>:let @/='\<'.expand("<cword>").'\>'<cr>
nnoremap <M-8> :call MyHigh()<cr>
func! MyHigh()
  let save = @/
  if len(save) != 0
    let @/ =""
    return
  endif
  let @/='\<'.expand("<cword>").'\>'
endf
nnoremap <S-MiddleMouse> <LeftMouse>:set ic<cr>:let @/='\<'.expand("<cword>").'\>'<cr>
map <M-MiddleMouse> 'a
map <M-h> :help 
nmap <silent> <F4> :TlistToggle<cr>
nmap <silent> <F12> <Plug>ToggleProject

" let g:indexer_disableCtagsWarning=1
" unmap <silent> <down>
"

" from D:\skill\vim\plugin\cscope_maps\cscope_maps.vim
nmap <C-\>s :cs find s <C-R>=expand("<cword>")<CR><CR>
nmap <C-\>g :cs find g <C-R>=expand("<cword>")<CR><CR>
nmap <C-\>c :cs find c <C-R>=expand("<cword>")<CR><CR>
nmap <C-\>t :cs find t <C-R>=expand("<cword>")<CR><CR>
nmap <C-\>e :cs find e <C-R>=expand("<cword>")<CR><CR>
nmap <C-\>f :cs find f <C-R>=expand("<cfile>")<CR><CR>
nmap <C-\>i :cs find i ^<C-R>=expand("<cfile>")<CR>$<CR>
nmap <C-\>d :cs find d <C-R>=expand("<cword>")<CR><CR>

nmap <C-@>s :scs find s <C-R>=expand("<cword>")<CR><CR>
nmap <C-@>g :scs find g <C-R>=expand("<cword>")<CR><CR>
nmap <C-@>c :scs find c <C-R>=expand("<cword>")<CR><CR>
nmap <C-@>t :scs find t <C-R>=expand("<cword>")<CR><CR>
nmap <C-@>e :scs find e <C-R>=expand("<cword>")<CR><CR>
nmap <C-@>f :scs find f <C-R>=expand("<cfile>")<CR><CR>
nmap <C-@>i :scs find i ^<C-R>=expand("<cfile>")<CR>$<CR>
nmap <C-@>d :scs find d <C-R>=expand("<cword>")<CR><CR>

nmap <C-@><C-@>s :vert scs find s <C-R>=expand("<cword>")<CR><CR>
nmap <C-@><C-@>g :vert scs find g <C-R>=expand("<cword>")<CR><CR>
nmap <C-@><C-@>c :vert scs find c <C-R>=expand("<cword>")<CR><CR>
nmap <C-@><C-@>t :vert scs find t <C-R>=expand("<cword>")<CR><CR>
nmap <C-@><C-@>e :vert scs find e <C-R>=expand("<cword>")<CR><CR>
nmap <C-@><C-@>f :vert scs find f <C-R>=expand("<cfile>")<CR><CR>
nmap <C-@><C-@>i :vert scs find i ^<C-R>=expand("<cfile>")<CR>$<CR>
nmap <C-@><C-@>d :vert scs find d <C-R>=expand("<cword>")<CR><CR>
" nnoremap <F5> :GundoToggle<CR>
" D:\skill\Apps\Calibre\calibre.exe
map <M-p> :pta <C-r>=expand("<cword>")<cr><cr>

" function! SuperCleverTab()
"   "check if at beginning of line or after a space
"   if strpart( getline('.'), 0, col('.')-1 ) =~ '^\s*$'
"     return "\<Tab>"
"   else
"     " do we have omni completion available
"     if &omnifunc != ''
"       "use omni-completion 1. priority
"       return "\<C-X>\<C-O>"
"     elseif &dictionary != ''
"       " no omni completion, try dictionary completio
"       return "\<C-X>\<C-K>"
"     else
"       "use omni completion or dictionary completion
"       "use known-word completion
"       return "\<C-X>\<C-N>"
"     endif
"   endif
" endfunction
" " bind function to the tab key
" inoremap <Tab> <C-R>=SuperCleverTab()<cr>

vnoremap <silent> <Enter> :EasyAlign<cr>

" for UltiSnips.txt
let g:UltiSnipsExpandTrigger="<tab>"
let g:UltiSnipsJumpForwardTrigger="<tab>"
let g:UltiSnipsJumpBackwardTrigger="<s-tab>"

map <C-tab> gt
map <C-S-tab> gT
imap <M-9> (

" nmap <silent> <M-1> :brewind<CR>
" nmap <silent> <M-2> :brewind \| 1bn<CR>
" nmap <silent> <M-3> :brewind \| 2bn<CR>
" nmap <silent> <M-4> :brewind \| 3bn<CR>
" nmap <silent> <M-5> :brewind \| 4bn<CR>
" nmap <silent> <M-6> :brewind \| 5bn<CR>
" nmap <silent> <M-7> :brewind \| 6bn<CR>
" nmap <silent> <M-8> :brewind \| 7bn<CR>
" nmap <silent> <M-9> :brewind \| 8bn<CR>
" nmap <silent> <M-0> :brewind \| 9bn<CR>
"nmap <silent> <unique> <C-W>o :call CloseAllBut()<cr>
"nmap <silent> <unique> <C-W><C-o> <C-W>o

function! CloseAllBut()
  let lzsave=&lz
  set lz
  let has_tag = 0
  "let has_nerdtree = 0
  "if exists('t:NERDTreeBufName')
    "let has_nerdtree = 1
  "endif
  if (bufwinnr("__Tag_List__") != -1)
    let has_tag = 1
  endif
  only
  "if (has_nerdtree == 1)
    "NERDTreeToggle
    "silent! wincmd p
  "endif
  if (has_tag == 1)
    TlistToggle
    silent! wincmd p
  endif

  let &lz=lzsave
  unlet lzsave
endfunction

"nnoremap <F8> :execute("NERDTreeToggle")<cr>
nnoremap <C-n> :tabnew<cr>
nnoremap <S-ScrollWheelUp> :zl<cr>
nnoremap <ScrollWheelUp> <C-E>
vmap <M-c> <C-c>

"unmap <C-Y>
map <M-j> <C-e>
map <M-k> <C-Y>
nnoremap <silent> <F2> :NumbersToggle<CR>
