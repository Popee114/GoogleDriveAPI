namespace DocumentHandler.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.IO.Compression;
    using System.Linq;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Threading;
    using System.Threading.Tasks;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;

    using DocumentHandler.Models;

    using Google.Apis.Auth.OAuth2;
    using Google.Apis.Auth.OAuth2.Flows;
    using Google.Apis.Auth.OAuth2.Responses;
    using Google.Apis.Docs.v1;
    using Google.Apis.Drive.v3;
    using Google.Apis.Drive.v3.Data;
    using Google.Apis.Services;

    using Microsoft.AspNetCore.Mvc;
    using Microsoft.EntityFrameworkCore;
    using Microsoft.Extensions.Logging;

    using Comment = DocumentFormat.OpenXml.Wordprocessing.Comment;
    using Document = DocumentHandler.Models.Document;
    using File = Google.Apis.Drive.v3.Data.File;

    public record DocumentRecord(long FileId, string DocumentId);

    [ApiController]
    [Route("api/v1")]
    public class DocumentHandlerController : ControllerBase
    {
        /// <summary>
        /// Логгер.
        /// </summary>
        private readonly ILogger<DocumentHandlerController> _logger;

        /// <summary>
        /// Id клиента.
        /// </summary>
        private readonly string CLIENT_ID = "483028773734-03ukk40g36s51u4otgni6h1jfmq4filk.apps.googleusercontent.com";

        /// <summary>
        /// Секрет клиента.
        /// </summary>
        private readonly string CLIENT_SECRET = "nDruP7QLsy8ESi60A4jl4s59";

        /// <summary>
        /// Креды.
        /// </summary>
        private readonly UserCredential Credential;

        /// <summary>
        /// База данных.
        /// </summary>
        private readonly DocumentContext Db = new DocumentContext();

        /// <summary>
        /// Драйв-сервис.
        /// </summary>
        private readonly DriveService _driveService;

        /// <summary>
        /// Токен обновления токена доступа.
        /// </summary>
        private readonly string REFRESH_TOKEN =
            "1//04qaAmhIaxa6ACgYIARAAGAQSNwF-L9IrRgV5UGs6nvHhqaQuzix2yUKZx4duVrP1TuBPj9eFXpet3GN0c4tjmcdqmlcI4eoErl4";

        /// <summary>
        /// Конструктор.
        /// </summary>
        /// <param name="logger">Логгер.</param>
        public DocumentHandlerController(ILogger<DocumentHandlerController> logger)
        {
            _logger = logger;
            Credential = GetUserCredential(CLIENT_ID, CLIENT_SECRET, REFRESH_TOKEN);
            _driveService = GetDriveService(Credential);
            //var fileId = "1q4EN3MoBn3DjGqc4Diqm2dJMJVYMIbd5tMMYea6KRA0";
            //var outputStream = new FileStream("test.docx", FileMode.Create);
            //DriveService.Files.Export(fileId, "application/vnd.openxmlformats-officedocument.wordprocessingml.document").Download(outputStream);
            //outputStream.Close();
        }

        /// <summary>
        /// Конец работы с документом.
        /// </summary>
        /// <param name="fileToSave">Файл, который следует сохранить..</param>
        [HttpPost("save/")]
        public async Task CloseDocument(DocumentRecord fileToSave)
        {
            var file = _driveService.Files.Get(fileToSave.DocumentId);
            var fileName = file.Execute().Name;

            if (IsFileEditable(fileName))
                await SaveDocument(fileToSave.FileId, file.FileId);

            _driveService.Files.Delete(fileToSave.DocumentId).Execute();
        }

        /// <summary>
        /// Открыть документ.
        /// </summary>
        /// <param name="fileId">Id документа.</param>
        /// <returns>Объект со ссылкой на документ.</returns>
        [HttpGet("open/")]
        public async Task<DocumentInfo> OpenDocumentAsync(long fileId)
        {
            //Скачивание документа из БД
            //var qwe = await Db.Documents.FirstOrDefaultAsync(x => x.Id == 3);
            //using (var fs = new FileStream(@"C:\Users\KazakovPA\Desktop\qqqqqqqqw.docx", FileMode.Create, FileAccess.Write))
            //{
            //    fs.Write(qwe.FileBytes, 0, qwe.FileBytes.Length);
            //}
            var downloadingFile = await Db.Documents.FindAsync(fileId);
            var documentId = await GetFileId(downloadingFile);
            downloadingFile.DocumentId = documentId;
            await Db.SaveChangesAsync();
            await GivePermission(documentId);
            var link = GetLink(documentId, downloadingFile.Name);
            return new DocumentInfo
            {
                Id = downloadingFile.Id,
                DocumentId = documentId,
                Link = link
            };
        }

        /// <summary>
        /// Метод отправки комментариев по документу.
        /// </summary>
        [HttpPost("reject/")]
        public async Task RejectDocument(DocumentRecord fileToReject)
        {
            var previousFileName = (string) Db.Documents.Find(fileToReject.FileId).Name.Clone();

            if (IsFileEditable(previousFileName))
            {
                await SaveDocument(fileToReject.FileId, fileToReject.DocumentId);

                var notes = await RetrieveComments(fileToReject.FileId);
                InsertComments(notes);
            }

            await _driveService.Files.Delete(fileToReject.DocumentId).ExecuteAsync();
        }

        /// <summary>
        /// Эксопрт файла из Драйва.
        /// </summary>
        /// <param name="fileId">Id файла.</param>
        private void ExportFile(string fileId)
        {
            var request = _driveService.Files.List();
            var response = request.Execute();
            var downloadFile = response.Files.First(file => file.Id == fileId);
            var getRequest = _driveService.Files.Get(downloadFile.Id);
            var fileStream = new FileStream($"Downloads/{downloadFile.Name}", FileMode.Create, FileAccess.ReadWrite);
            getRequest.Download(fileStream);
            fileStream.Close();
        }

        private DocsService GetDocsService(UserCredential credential)
        {
            return new DocsService(
                new BaseClientService.Initializer
                {
                    HttpClientInitializer = credential
                });
        }

        /// <summary>
        /// Инициализация Драйв-сервиса по кредам.
        /// </summary>
        /// <param name="credential">Креды.</param>
        /// <returns>Экземпляр Драйв-сервиса.</returns>
        private DriveService GetDriveService(UserCredential credential)
        {
            return new DriveService(
                new BaseClientService.Initializer
                {
                    HttpClientInitializer = credential
                });
        }

        private async Task<CommentList> GetFileComments(string documentId)
        {
            var commentsRequest = _driveService.Comments.List(documentId);
            commentsRequest.Fields = "*";
            return await commentsRequest.ExecuteAsync();
        }

        private async Task<string> GetFileId(Document file)
        {
            try
            {
                var fileFromService = await _driveService.Files.Get(file.DocumentId).ExecuteAsync();
                return fileFromService.Id;
            }
            catch (Exception e)
            {
                return await UploadFileToDrive(file, false);
            }
        }

        private string GetLink(string fileId, string fileName)
        {
            if (IsFileEditable(fileName))
                return $"https://docs.google.com/document/d/{fileId}/edit?rm=demo";

            return $"https://drive.google.com/file/d/{fileId}/preview";
        }

        /// <summary>
        /// Инициализация кредов.
        /// </summary>
        /// <param name="clientId">Id клиента.</param>
        /// <param name="clientSecret">Секрет клиента.</param>
        /// <param name="refreshToken">Токен обновления токена доступа.</param>
        /// <returns></returns>
        private UserCredential GetUserCredential(string clientId, string clientSecret, string refreshToken)
        {
            var secrets = new ClientSecrets
            {
                ClientId = clientId,
                ClientSecret = clientSecret
            };

            var token = new TokenResponse { RefreshToken = refreshToken };
            return new UserCredential(
                new GoogleAuthorizationCodeFlow(
                    new GoogleAuthorizationCodeFlow.Initializer
                    {
                        ClientSecrets = secrets
                    }),
                Environment.UserName, token);
        }

        /// <summary>
        /// Задание прав документу.
        /// </summary>
        /// <param name="fileId">Id файла.</param>
        private async Task GivePermission(string fileId)
        {
            await _driveService.Permissions.Create(new Permission
                {
                    Role = "commenter",
                    Type = "anyone"
                },
                fileId).ExecuteAsync();
        }

        private bool IsFileEditable(string fileName)
        {
            return fileName.EndsWith(".doc") || fileName.EndsWith(".docx");
        }

        /// <summary>
        ///     Выбор комментариев и текста, предложенного к вставке, с ключами.
        /// </summary>
        /// <param name="fileId"> Идентификатор файла. </param>
        /// <returns> Списко комментариев с ключами. </returns>
        private async Task<List<string>> RetrieveComments(long fileId)
        {
            var commentsData = new Dictionary<string, string>();

            var downloadingFile = await Db.Documents.FindAsync(fileId);

            await using (var stream = new MemoryStream(downloadingFile.FileBytes))
            {
                using (var wordDoc = WordprocessingDocument.Open(stream, false))
                {
                    // Поиск комментариев и текста, в котором они приписаны.
                    var comments = wordDoc.MainDocumentPart?.WordprocessingCommentsPart?.Comments.Elements<Comment>().ToList();
                    var commentRangeStarts = wordDoc.MainDocumentPart.Document.Body.Descendants<SdtElement>().SelectMany(x => x.Descendants<CommentRangeStart>()).ToList();

                    // Выбор комментария с соответствующим текстом.
                    for (var i = 0; i < comments.Count; i++)
                    {
                        var keyComment = wordDoc.MainDocumentPart.Document.Body.Descendants<Paragraph>()
                            .Where(x => x.Descendants<CommentRangeStart>().Contains(commentRangeStarts[i]))
                            .SelectMany(x => x.Descendants<Text>().Select(x => x.Text))
                            .ToList();

                        var key = string.Join(" ", keyComment.ToArray());
                        commentsData.Add(key, comments[i].InnerText);
                    }

                    // Поиск текста, предложенного для вставки, и текста, к которому он относится.
                    var runs = wordDoc.MainDocumentPart.Document.Body.Descendants<SdtElement>().SelectMany(x => x.Descendants<Run>()).ToList();
                    var tags = wordDoc.MainDocumentPart.Document.Body.Descendants<SdtElement>()
                        .SelectMany(x => x.Descendants<Run>())
                        .SelectMany(x => x.Parent.Parent.Parent.ChildElements.GetItem(0).Descendants<Tag>())
                        .ToList();

                    // Выборка предложенного текста с соответствующим текстом.
                    for (var i = 0; i < runs.Count; i++)
                    {
                        var keyAndComment = wordDoc.MainDocumentPart.Document.Body.Descendants<Paragraph>()
                            .Where(x => x.Descendants<Tag>().Contains(tags[i]) && x.Descendants<Run>().Contains(runs[i]))
                            .SelectMany(x => x.Descendants<Text>())
                            .ToList();
                        if (keyAndComment.Count != 0)
                            commentsData.Add(keyAndComment[0].Text, keyAndComment[1].Text);
                    }
                }
            }

            var notes = new List<string>();
            foreach (var (key, value) in commentsData)
            {
                var note = key + " - " + value;
                notes.Add(note);
            }

            return notes;
        }

        private async Task SaveDocument(long fileId, string documentId)
        {
            var file = _driveService.Files.Get(documentId);
            Thread.Sleep(3000);
            byte[] fileBytes;
            await using (var stream = new MemoryStream())
            {
                file.Download(stream);
                fileBytes = stream.GetBuffer();
            }

            var currentFile = await Db.Documents.FirstOrDefaultAsync(doc => doc.Id == fileId);
            currentFile.FileBytes = fileBytes;
            await Db.SaveChangesAsync();
        }

        /// <summary>
        ///     Заполнение БД документами.
        /// </summary>
        private async Task UploadFilesFrom()
        {
            var pathFolder = "Files to upload";
            var dirs = Directory.GetFiles(pathFolder);
            var files = new List<FileInfo>();
            foreach (var file in dirs)
            {
                files.Add(new FileInfo(file));
            }

            foreach (var file in files)
            {
                await UploadFileToDB(file);
            }
        }

        private void UploadFilesFromBDtoDrive()
        {
            var documents = Db.Documents.ToList();
            foreach (var document in documents)
            {
                UploadFileToDrive(document, false);
            }
        }

        private async Task UploadFileToDB(FileInfo file)
        {
            using (var stream = file.OpenRead())
            {
                var length = stream.Length;
                var bytes = new byte[length];
                var position = 0;
                int reader;
                while ((reader = await stream.ReadAsync(bytes, position, (int) (length - position))) > 0)
                    position += reader;
                Db.Documents.Add(new Document { Name = file.Name, FileBytes = bytes });
                await Db.SaveChangesAsync();
            }
        }

        /// <summary>
        /// Загрузка файла на Драйв.
        /// </summary>
        /// <param name="fileName">Id файла.</param>
        /// <returns>Id загруженного файла.</returns>
        private async Task<string>  UploadFileToDrive(Document file, bool isRejected)
        {
            FilesResource.CreateMediaUpload request;

            var mimeType = isRejected ? "application/vnd.google-apps.document" : "application/octet-stream";
            var fileName = isRejected ? $"[Отказ] {file.Name}" : file.Name;

            var fileMetaData = new File { Name = fileName, MimeType = mimeType };

            using (var stream = new MemoryStream(file.FileBytes))
            {
                request = _driveService.Files.Create(fileMetaData, stream, "application/octet-stream");
                await request.UploadAsync();
            }

            var fileToUpload = request.ResponseBody;
            return fileToUpload.Id;
        }

        private void InsertComments(List<string> notes)
        {
            byte[] byteArray = System.IO.File.ReadAllBytes(@"Assets/Rejecting letter template.docx");
            using (var stream = new MemoryStream())
            {
                stream.Write(byteArray, 0, byteArray.Length);

                using (var wordDoc = WordprocessingDocument.Open(stream, true))
                {
                    var docText = "";
                    using (var sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                    {
                        docText = sr.ReadToEnd();
                    }

                    var marks = new Dictionary<string, string>()
                    {
                        { "{{dateFrom}}", DateTime.Now.ToShortDateString()},
                        { "{{number}}", "Ы-99-00-212"},
                        { "{{organizationName}}", "OOO \"Организация\""},
                        { "{{organizationLeader}}", "А. Б. Сидоров"},
                        { "{{organizationAddress}}", "Какой-то там длиный адрес, типа индекс, город, улица и т.п."},
                        { "{{copyOrganizationName}}", "OOO \"Еще одна организация\""},
                        { "{{copyOrganizationLeader}}", "В. Г. Петров"},
                        { "{{copyOrganizationAddress}}", "Еще один какой-то там длиный адрес, типа индекс, город, улица и т.п., но уже другой"},
                        { "{{passportTargets}}", "Сахарная пудра"},
                        { "{{standartNumber}}", "0803-01"}
                    };

                    var body = wordDoc.MainDocumentPart.Document.Body;
                    var paragraphs = body.Descendants<Paragraph>().Where(x => Regex.IsMatch(x.InnerText, @".*\{{\w+\}}.*")).ToList();

                    foreach (Paragraph paragraph in paragraphs)
                    {
                        foreach (Match markMatch in Regex.Matches(paragraph.InnerText, @"\{{\w+\}}", RegexOptions.Compiled))
                        {
                            var paragraphMarkValue = markMatch.Value.Trim();
                            if (marks.TryGetValue(paragraphMarkValue, out var markValueFromCollection))
                            {
                                string editedParagraphText = paragraph.InnerText.Replace(markMatch.Value, markValueFromCollection);
                                var first = paragraph.Elements<Run>().Last();
                                var props = first.RunProperties;

                                var formattedRun = new Run();
                                var runPro = new RunProperties();
                                runPro.Append(new Bold(), new Text(editedParagraphText));
                                runPro.Color = (Color)props.Color?.Clone();
                                runPro.RunFonts = (RunFonts)props.RunFonts?.Clone();
                                runPro.Bold = (Bold)props.Bold?.Clone();
                                formattedRun.Append(runPro);

                                paragraph.RemoveAllChildren<Run>();
                                paragraph.AppendChild(formattedRun);
                            }
                        }
                    }

                    var paragraphNode = wordDoc.MainDocumentPart.Document.Body.Descendants<Paragraph>()
                        .First(x => x.InnerText == "{{rejectingNotes}}");
                    paragraphNode.RemoveAllChildren<Run>();

                    var newParagraphs = new List<Paragraph>();

                    foreach (var note in notes)
                    {
                        var paragraph = new Paragraph { ParagraphProperties = (ParagraphProperties?)paragraphNode.ParagraphProperties.CloneNode(true) };
                        var run = new Run(new RunProperties { RunFonts = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" } }, new Text($"- {note};"));
                        paragraph.AppendChild(run);
                        newParagraphs.Add(paragraph);
                    }

                    newParagraphs.Reverse();

                    foreach (var newParagraph in newParagraphs)
                        paragraphNode.InsertAfterSelf(newParagraph);

                    using (var sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                    {
                        sw.Write(docText);
                    }
                }

                System.IO.File.WriteAllBytes(@"Assets/Rejecting letter.docx", stream.ToArray());
            }
        }
    }
}