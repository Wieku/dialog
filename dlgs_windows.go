package dialog

import (
	"fmt"
	"github.com/TheTitanrain/w32"
	"github.com/rodrigocfd/windigo/win/com/com"
	"github.com/rodrigocfd/windigo/win/com/com/comco"
	"github.com/rodrigocfd/windigo/win/com/shell"
	"github.com/rodrigocfd/windigo/win/com/shell/shellco"
	"strings"
)

func init() {
	w32.CoInitializeEx(w32.COINIT_MULTITHREADED)
}

type WinDlgError int

func (e WinDlgError) Error() string {
	return fmt.Sprintf("CommDlgExtendedError: %#x", e)
}

func (b *MsgBuilder) yesNo() bool {
	r := w32.MessageBox(w32.HWND(0), b.Msg, firstOf(b.Dlg.Title, "Confirm?"), w32.MB_YESNO|w32.MB_ICONQUESTION)
	return r == w32.IDYES
}

func (b *MsgBuilder) errorYesNo() bool {
	r := w32.MessageBox(w32.HWND(0), b.Msg, firstOf(b.Dlg.Title, "Error"), w32.MB_YESNO|w32.MB_ICONERROR)
	return r == w32.IDYES
}

func (b *MsgBuilder) info() {
	w32.MessageBox(w32.HWND(0), b.Msg, firstOf(b.Dlg.Title, "Information"), w32.MB_OK|w32.MB_ICONINFORMATION)
}

func (b *MsgBuilder) error() {
	w32.MessageBox(w32.HWND(0), b.Msg, firstOf(b.Dlg.Title, "Error"), w32.MB_OK|w32.MB_ICONERROR)
}

type filedlg struct {
	dlg    shell.IFileDialog
	pathSH shell.IShellItem
}

func (d filedlg) Release() {
	if d.pathSH != nil {
		d.pathSH.Release()
	}

	if d.dlg != nil {
		d.dlg.Release()
	}
}

func (b *FileBuilder) load() (string, error) {
	d := openfile(b, 0)

	defer d.Release()

	if !d.dlg.Show(0) {
		return "", ErrCancelled
	}

	return d.dlg.GetResultDisplayName(shellco.SIGDN_DESKTOPABSOLUTEEDITING), nil
}

func (b *FileBuilder) loadMultiple() ([]string, error) {
	d := openfile(b, shellco.FOS_ALLOWMULTISELECT)

	defer d.Release()

	if !d.dlg.Show(0) {
		return nil, ErrCancelled
	}

	sDlg, _ := d.dlg.(shell.IFileOpenDialog)

	return sDlg.GetResults().ListDisplayNames(shellco.SIGDN_DESKTOPABSOLUTEEDITING), nil
}

func (b *FileBuilder) save() (string, error) {
	d := savefile(b)

	defer d.Release()

	if !d.dlg.Show(0) {
		return "", ErrCancelled
	}

	return d.dlg.GetResultDisplayName(shellco.SIGDN_DESKTOPABSOLUTEEDITING), nil
}

func openfile(b *FileBuilder, mFlag shellco.FOS) (d filedlg) {
	d.dlg = shell.NewIFileOpenDialog(com.CoCreateInstance(
		shellco.CLSID_FileOpenDialog, nil,
		comco.CLSCTX_INPROC_SERVER,
		shellco.IID_IFileOpenDialog))

	common(b, d.dlg, d, shellco.FOS_FILEMUSTEXIST|mFlag)

	return d
}

func savefile(b *FileBuilder) (d filedlg) {
	d.dlg = shell.NewIFileSaveDialog(com.CoCreateInstance(
		shellco.CLSID_FileSaveDialog, nil,
		comco.CLSCTX_INPROC_SERVER,
		shellco.IID_IFileSaveDialog))

	common(b, d.dlg, d, shellco.FOS_OVERWRITEPROMPT)

	return d
}

func common(b *FileBuilder, dlg shell.IFileDialog, fdlg filedlg, flags shellco.FOS) {
	dlg.SetOptions(dlg.GetOptions() | shellco.FOS_NOCHANGEDIR | flags)

	if b.Filters != nil && len(b.Filters) > 0 {
		var fSpec []shell.FilterSpec

		for _, filt := range b.Filters {
			fSpec = append(fSpec, shell.FilterSpec{
				Name: filt.Desc,
				Spec: "*." + strings.Join(filt.Extensions, ";*."),
			})
		}

		dlg.SetFileTypes(fSpec)
	}

	if b.StartDir != "" {
		fdlg.pathSH, _ = shell.NewShellItemFromPath(b.StartDir)

		dlg.SetFolder(fdlg.pathSH)
	}

	dlg.SetTitle(b.Dlg.Title)
}

func (b *DirectoryBuilder) browse() (string, error) {
	d := openfile(&FileBuilder{
		Dlg:      b.Dlg,
		StartDir: b.StartDir,
		Filters:  nil,
	}, shellco.FOS_PICKFOLDERS|shellco.FOS_PATHMUSTEXIST)

	defer d.Release()

	if !d.dlg.Show(0) {
		return "", ErrCancelled
	}

	return d.dlg.GetResultDisplayName(shellco.SIGDN_DESKTOPABSOLUTEEDITING), nil
}
