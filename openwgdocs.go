package main

import (
	"bufio"
	"context"
	"crypto/md5"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"io"
	"log"
	"mime"
	"net/http"
	"net/url"
	"os"
	"os/exec"
	"path"
	"path/filepath"
	"runtime"
	"strings"
	"time"

	"golang.org/x/oauth2"
	"golang.org/x/sys/windows/registry"
	"google.golang.org/api/drive/v2"
	"google.golang.org/api/option"
)

func getConfig() *oauth2.Config {
	return &oauth2.Config{
		ClientID:    "839856140400-of8r47k6ieg5bsrup0e23oaa9mel2kef.apps.googleusercontent.com",
		Scopes:      []string{"https://www.googleapis.com/auth/drive"},
		RedirectURL: "http://localhost:3333/update_token",
		Endpoint: oauth2.Endpoint{
			AuthURL:  "https://accounts.google.com/o/oauth2/auth",
			TokenURL: "https://accounts.google.com/o/oauth2/token",
		},
	}
}

func openInBrowser(url string) {
	var err error

	switch runtime.GOOS {
	case "linux":
		err = exec.Command("xdg-open", url).Start()
	case "windows":
		err = exec.Command("rundll32", "url.dll,FileProtocolHandler", url).Start()
	case "darwin":
		err = exec.Command("open", url).Start()
	default:
		err = fmt.Errorf("unsupported platform")
	}
	if err != nil {
		fmt.Println("Please open the following url in your browser {}", url)
	}
}

func waitForToken() *url.Values {
	paramsCh := make(chan url.Values)
	srv := &http.Server{Addr: ":3333"}

	http.HandleFunc("/update_token", func(w http.ResponseWriter, r *http.Request) {
		if len(r.URL.Query()) == 0 {
			io.WriteString(w, "<script>location.replace(location.href.replace('#', '?'))</script>")
		} else {
			io.WriteString(w, "Close this page and return to program!")
			paramsCh <- r.URL.Query()
		}
	})
	go srv.ListenAndServe()

	params := <-paramsCh
	close(paramsCh)
	srv.Shutdown(context.TODO())

	return &params
}

func openPrompt(config *oauth2.Config) {
	openInBrowser(config.AuthCodeURL("state-1234", oauth2.SetAuthURLParam("response_type", "token")))
}

func buildToken(params *url.Values) *oauth2.Token {
	return &oauth2.Token{
		AccessToken: params.Get("access_token"),
		TokenType:   "Bearer",
		Expiry:      time.Now().Add(1000000000 * 3599),
	}
}

func updateToken(config *oauth2.Config) *oauth2.Token {
	openPrompt(config)
	params := waitForToken()
	return buildToken(params)
}

func tokenFromFile(file string) (*oauth2.Token, error) {
	f, err := os.Open(file)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	tok := &oauth2.Token{}
	err = json.NewDecoder(f).Decode(tok)
	if !tok.Valid() {
		return nil, fmt.Errorf("token expired")
	}
	return tok, err
}

func saveToken(path string, token *oauth2.Token) {
	f, err := os.OpenFile(path, os.O_RDWR|os.O_CREATE|os.O_TRUNC, 0600)
	if err != nil {
		log.Fatalf("Unable to cache oauth token: %v", err)
	}
	defer f.Close()
	json.NewEncoder(f).Encode(token)
}

func getClient() *http.Client {
	config := getConfig()
	tokFile := "token.json"
	tok, err := tokenFromFile(path.Join(getExePath(), tokFile))
	if err != nil {
		tok = updateToken(config)
		saveToken(path.Join(getExePath(), tokFile), tok)
	}
	return config.Client(context.Background(), tok)
}

func fileHash(filename string) string {
	file, err := os.Open(filename)
	if err != nil {
		log.Fatalln(err)
	}
	defer file.Close()

	hash := md5.New()
	if _, err := io.Copy(hash, file); err != nil {
		log.Fatal(err)
	}
	return hex.EncodeToString(hash.Sum(nil))
}

func findFileByHash(hash string) string {
	file, err := os.OpenFile(path.Join(getExePath(), "filelist"), os.O_RDONLY|os.O_CREATE, 0600)
	if err != nil {
		log.Fatalln(err)
	}
	defer file.Close()

	scanner := bufio.NewScanner(file)
	for scanner.Scan() {
		line := scanner.Text()
		fHash := line[:32]
		if fHash == hash {
			return line[32:]
		}
	}

	return ""
}

func findFile(filename string) string {
	return findFileByHash(fileHash(filename))
}

func appName(fileExt string) string {
	switch strings.ToLower(fileExt) {
	case ".doc", ".docx":
		return "document"
	case ".xls", ".xlsx":
		return "spreadsheets"
	case ".ppt", ".pptx":
		return "presentation"
	}
	return ""
}

func pause() {
	fmt.Scanln()
}

func getFilenameFromArgs() string {
	if len(os.Args) != 2 {
		return ""
	}
	return os.Args[1]
}

func buildFileURL(fileId, appName string) string {
	return fmt.Sprintf("https://docs.google.com/%s/d/%s/edit", appName, fileId)
}

func openFileFromCache(filename string) bool {
	if fileId := findFile(filename); fileId != "" {
		if appName := appName(filepath.Ext(filename)); appName != "" {
			openInBrowser(buildFileURL(fileId, appName))
			return true
		} else {
			fmt.Println("Can't determine filetype. Press enter to exit.")
			pause()
			return true
		}
	}
	return false
}

func saveFileToCache(fileHash, fileId string) {
	f, err := os.OpenFile(path.Join(getExePath(), "filelist"), os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0600)
	if err != nil {
		log.Println(err)
	}
	defer f.Close()
	if _, err := f.WriteString(fileHash + fileId + "\n"); err != nil {
		log.Println(err)
	}
}

func uploadAndOpenFile(filename string) error {
	appName := appName(filepath.Ext(filename))
	if appName == "" {
		return fmt.Errorf("can't determine filetype")
	}
	ctx := context.Background()
	client := getClient()
	srv, err := drive.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return err
	}

	file, err := os.Open(filename)
	if err != nil {
		return err
	}
	defer file.Close()

	f := &drive.File{
		Title:    filepath.Base(filename),
		MimeType: mime.TypeByExtension(filepath.Ext(filename)),
	}
	res, err := srv.Files.
		Insert(f).
		Media(file).
		ProgressUpdater(func(now, size int64) { fmt.Printf("%d%%", now/size) }).
		Do()
	if err != nil {
		return err
	}

	openInBrowser(buildFileURL(res.Id, appName))
	saveFileToCache(fileHash(filename), res.Id)
	return nil
}

func getExePath() string {
	ex, err := os.Executable()
	if err != nil {
		panic(err)
	}
	return filepath.Dir(ex)
}

func getExeName() string {
	ex, err := os.Executable()
	if err != nil {
		panic(err)
	}
	return filepath.Base(ex)
}

func amAdmin() bool {
	_, err := os.Open("\\\\.\\PHYSICALDRIVE0")
	return err == nil
}

func addToOpenWith(ext string) {
	k, _, err := registry.CreateKey(
		registry.CURRENT_USER,
		fmt.Sprintf(`Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\%s\OpenWithList`, ext),
		registry.QUERY_VALUE|registry.SET_VALUE,
	)
	if err != nil {
		log.Fatal(err)
	}
	defer k.Close()

	mruList, _, err := k.GetStringValue("MRUList")
	if err != nil {
		mruList = ""
	}
	max := rune('a' - 1)
	curExeName := getExeName()
	for _, char := range mruList {
		if char > max {
			max = char
		}
		program, _, err := k.GetStringValue(string(char))
		if err != nil {
			panic(err)
		}
		if program == curExeName {
			return
		}
	}

	max += 1
	k.SetStringValue(string(max), curExeName)
	k.SetStringValue("MRUList", mruList+string(max))
}

func associateFileExts() {
	curExeName := getExeName()
	k, _, err := registry.CreateKey(
		registry.CLASSES_ROOT,
		fmt.Sprintf(`Applications\%s\shell\open\command`, curExeName),
		registry.QUERY_VALUE|registry.SET_VALUE,
	)
	if err != nil {
		panic(err)
	}
	k.SetStringValue(
		"",
		fmt.Sprintf(`"%s\%s" "%%1"`, getExePath(), getExeName()),
	)

	k, _, err = registry.CreateKey(
		registry.CLASSES_ROOT,
		fmt.Sprintf(`Applications\%s\SupportedTypes`, curExeName),
		registry.QUERY_VALUE|registry.SET_VALUE,
	)
	if err != nil {
		panic(err)
	}

	types := []string{".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx"}

	for _, ext := range types {
		err = k.SetStringValue(ext, "")
		if err != nil {
			panic(err)
		}
	}

	for _, ext := range types {
		addToOpenWith(ext)
	}
}

func setup() bool {
	if runtime.GOOS != "windows" {
		fmt.Println("Setup is only supported on Windows, exiting. Press enter to exit.")
		return false
	}

	if !amAdmin() {
		fmt.Println("Run as admin to run setup! Press enter to exit.")
		return false
	}

	associateFileExts()
	return true
}

func main() {
	filename := getFilenameFromArgs()
	if filename == "" {
		if setup() {
			fmt.Println("Setup complete. Use \"Open with\" option on document files to use the program. Press enter to exit.")
		}
		pause()
		return
	}
	opened := openFileFromCache(filename)
	if opened {
		return
	}
	error := uploadAndOpenFile(filename)
	if error != nil {
		fmt.Println(error)
		fmt.Println("Press enter to exit.")
		pause()
	}
}
