using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using UnityEngine;
using UnityEngine.UI;
using ZXing;
using System.Threading.Tasks;
using TMPro;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

public class QRCodeScanner : MonoBehaviour
{
    public RawImage cameraPreview;
    public RenderTexture cameraTexture;
    public TextMeshProUGUI nameText, sexText, title1Text, title2Text, unitText, countText, tableText;
    public GameObject namePanel, titlePanel, unitPanel, tablePanel;
    public GameObject nameLabel, titleLabel;
    public GameObject standbyScreen;
    public TextMeshProUGUI tmpgui;
    public TMP_InputField manualInputField;
    public int scansPerSecond = 5;

    private WebCamTexture webcamTexture;
    private IBarcodeReader barcodeReader;

    private Dictionary<string, ExcelRow> guestDict = new Dictionary<string, ExcelRow>();
    private HashSet<string> checkedIn = new HashSet<string>();

    private string dataPath;
    private float standbyTimer;
    public AudioClip qrSound;
    public AudioSource audioSource;
    public float standbyTimeout = 60f;
    public float fadeDuration = 1f; // Thời gian fade-in chữ

    // NPOI objects
    private IWorkbook workbook;
    private ISheet worksheet;

    public LogManager logManager;
    // private Animator nameAnimator, titleAnimator, unitAnimator, tableAnimator;

    void Start()
    {
        if (Display.displays.Length > 1)
        {
            Display.displays[1].Activate();
        }

        string filePath = Path.Combine(Path.GetDirectoryName(Application.dataPath), "Data");
        dataPath = Path.Combine(filePath, "Data.xlsx");

        LoadExcel();
        InitWebcam();
        barcodeReader = new BarcodeReader();

        standbyScreen.SetActive(false);
        StartCoroutine(ScanRoutine());

        // Gắn sự kiện Enter cho input
        manualInputField.onSubmit.AddListener(delegate { ManualCheckIn(); ExitStandby(); });
        // nameAnimator = namePanel.GetComponent<Animator>();
        // titleAnimator = titlePanel.GetComponent<Animator>();
        // unitAnimator = unitPanel.GetComponent<Animator>();
        // tableAnimator = tablePanel.GetComponent<Animator>();
    }

    void Update()
    {
        // Tăng timer standby
        standbyTimer += Time.deltaTime;

        // Nếu quá timeout và chưa standby thì bật standby
        if (standbyTimer >= standbyTimeout && !standbyScreen.activeSelf)
        {
            EnterStandby();
        }

        // Nếu đang standby và người dùng bấm phím thì thoát standby
        // if (standbyScreen.activeSelf && Input.anyKeyDown)
        // {
        //     ExitStandby();
        // }

        // Ngoài ra, cho phép bấm Enter để nhập manual khi đang active input
        // if (manualInputField.isFocused && Input.GetKeyDown(KeyCode.Return))
        if (Input.GetKeyDown(KeyCode.Return))
        {
            ManualCheckIn();
            ExitStandby();
        }
    }

    void EnterStandby()
    {
        standbyScreen.SetActive(true);
        // if (webcamTexture != null && webcamTexture.isPlaying)
        // {
        //     webcamTexture.Pause(); // Tạm dừng camera để tiết kiệm tài nguyên
        // }
        // tmpgui.text = "Đang ở chế độ chờ...";
    }

    void ExitStandby()
    {
        standbyTimer = 0f;
        standbyScreen.SetActive(false);
        // if (webcamTexture != null && !webcamTexture.isPlaying)
        // {
        //     webcamTexture.Play();
        // }
        tmpgui.text = "Camera OK - Đang quét...";
    }

    IEnumerator ScanRoutine()
    {
        while (true)
        {
            // Nếu standby thì bỏ qua vòng quét để tránh lag
            // if (standbyScreen.activeSelf)
            // {
            //     yield return new WaitForSeconds(0.5f);
            //     continue;
            // }

            if (webcamTexture != null && webcamTexture.isPlaying)
            {
                // tmpgui.text = "Camera OK - Đang quét...";
                bool scanned = false;

                try
                {
                    var data = webcamTexture.GetPixels32();
                    var width = webcamTexture.width;
                    var height = webcamTexture.height;

                    var result = barcodeReader.Decode(data, width, height);
                    if (result != null)
                    {
                        standbyTimer = 0f; // reset timer khi có scan
                        ProcessQRCode(result.Text.Trim());
                        standbyScreen.SetActive(false);
                        scanned = true;
                        // tmpgui.text = $"Đã quét ID: {result.Text.Trim()}";
                        logManager?.AddLog($"Đã quét QR thành công: {result.Text.Trim()}");
                    }
                }
                catch (Exception e)
                {
                    Debug.LogWarning("Scan error: " + e.Message);
                    tmpgui.text = " Lỗi khi quét QR.";
                }

                if (scanned)
                {
                    yield return new WaitForSeconds(2f);
                }
            }
            else
            {
                tmpgui.text = " Không tìm thấy camera.";
            }

            yield return new WaitForSeconds(1f);
        }
    }

    public void ManualCheckIn()
    {
        string inputId = manualInputField.text.Trim();
        // uppercase first letter
        if (!string.IsNullOrEmpty(inputId))
        {
            inputId = char.ToUpper(inputId[0]) + inputId.Substring(1).ToLower();
        }
        if (!string.IsNullOrEmpty(inputId))
        {
            logManager?.AddLog($"Check-in thủ công: {inputId}");
            tmpgui.text = $"Check-in thủ công: {inputId}";
            ProcessQRCode(inputId);
            manualInputField.text = ""; // Xóa input sau khi nhập
        }
        else
        {
            // tmpgui.text = " Vui lòng nhập ID.";
        }
        standbyTimer = 0f; // reset standby
    }


    void InitWebcam()
    {
        WebCamDevice[] devices = WebCamTexture.devices;
        if (devices.Length > 0)
        {
            webcamTexture = new WebCamTexture(devices[0].name, cameraTexture.width, cameraTexture.height);
            cameraPreview.texture = webcamTexture;
            webcamTexture.Play();
            // cameraPreview.rectTransform.sizeDelta = new Vector2(webcamTexture.width, webcamTexture.height);
            cameraPreview.enabled = true;
            tmpgui.text = " Camera đã khởi động.";
        }
        else
        {
            tmpgui.text = " Không tìm thấy thiết bị camera.";
            Debug.LogError("No webcam found!");
        }
    }

    void LoadExcel()
    {
        using (FileStream file = new FileStream(dataPath, FileMode.Open, FileAccess.Read))
        {
            workbook = new XSSFWorkbook(file);
            worksheet = workbook.GetSheet("DS Final");
        }

        guestDict.Clear();

        int row = 1; // NPOI index bắt đầu từ 0, dữ liệu bạn ở hàng 2 => index = 1
        while (true)
        {
            IRow excelRow = worksheet.GetRow(row);
            if (excelRow == null) break;

            string id = excelRow.GetCell(1)?.ToString(); // Cột B => index 1
            if (string.IsNullOrEmpty(id)) break;

            guestDict[id] = new ExcelRow
            {
                RowNumber = row,
                ID = id,
                Sex = excelRow.GetCell(3)?.ToString(),
                Name = excelRow.GetCell(4)?.ToString(),
                Unit = excelRow.GetCell(5)?.ToString(),
                Title1 = excelRow.GetCell(6)?.ToString(),
                Title2 = excelRow.GetCell(7)?.ToString(),
                Table = excelRow.GetCell(11)?.ToString()
            };
            row++;
        }

        tmpgui.text = $"Đã tải dữ liệu từ Excel ({guestDict.Count} người).";
    }

    void ProcessQRCode(string id)
    {
        if (string.IsNullOrEmpty(id))
        {
            tmpgui.text = " ID không hợp lệ.";
            return;
        }
        if (!guestDict.ContainsKey(id))
        {
            tmpgui.text = $" ID không tồn tại: {id}";
            logManager?.AddLog($" Không tìm thấy ID: {id}");
            return;
        }
        ExcelRow guest = guestDict[id];
        IRow excelRow = worksheet.GetRow(guest.RowNumber);

        if (checkedIn.Contains(id))
        {
            tmpgui.text = $" Đã check-in trước đó: {id}";
            logManager?.AddLog($" Lặp lại ID đã check-in: {id}");
            checkedIn.Remove(id);
            // return;
        }
        else
        {
            if (excelRow.GetCell(15) == null) excelRow.CreateCell(15);
            excelRow.GetCell(15).SetCellValue(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
        }

        ShowInfo(guest);



        checkedIn.Add(id);

        SaveExcelAsync();
        countText.text = $"Số người đã checkin: {checkedIn.Count}";
        tmpgui.text = $"Check-in thành công: {guest.Name}";
        logManager?.AddLog($"Check-in thành công: {guest.Name} - ID: {id}");

        // Gọi hiệu ứng
        PlayScanEffect();
    }
    void PlayScanEffect()
    {
        // Phát âm thanh
        if (audioSource != null && qrSound != null)
        {
            audioSource.PlayOneShot(qrSound);
        }

        // Fade-in text thông tin khách
        // StartCoroutine(FadeInText(nameText));
        // StartCoroutine(FadeInText(sexText));
        // StartCoroutine(FadeInText(title1Text));
        // StartCoroutine(FadeInText(title2Text));
        // StartCoroutine(FadeInText(unitText));
        // StartCoroutine(FadeInText(tableText));
        // zoom in text panels
        StartCoroutine(ZoomInTextPanel(namePanel));
        StartCoroutine(ZoomInTextPanel(titlePanel));
        StartCoroutine(ZoomInTextPanel(unitPanel));
        StartCoroutine(ZoomInTextPanel(tablePanel));
        // get animator components and play zoom in animation

        // nameAnimator.CrossFade("NameAnim", 0.1f);
        // titleAnimator.CrossFade("TitleAnim", 0.1f);
        // unitAnimator.CrossFade("UnitAnim", 0.1f);
        // tableAnimator.CrossFade("TableAnim", 0.1f);
    }

    IEnumerator ZoomInTextPanel(GameObject panel)
    {
        if (panel == null) yield break;

        Vector3 originalScale = panel.transform.localScale;
        panel.transform.localScale = Vector3.zero;

        float elapsed = 0f;
        while (elapsed < fadeDuration)
        {
            elapsed += Time.deltaTime;
            float t = Mathf.Clamp01(elapsed / fadeDuration);
            panel.transform.localScale = Vector3.Lerp(Vector3.zero, originalScale, t);
            yield return null;
        }
    }

    // IEnumerator FadeInText(TextMeshProUGUI tmp)
    // {
    //     if (tmp == null) yield break;

    //     Color c = tmp.color;
    //     c.a = 0;
    //     tmp.color = c;

    //     float elapsed = 0f;
    //     while (elapsed < fadeDuration)
    //     {
    //         elapsed += Time.deltaTime;
    //         c.a = Mathf.Clamp01(elapsed / fadeDuration);
    //         tmp.color = c;
    //         yield return null;
    //     }
    // }
    async void SaveExcelAsync()
    {
        await Task.Run(() =>
        {
            using (FileStream file = new FileStream(dataPath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(file);
            }
        });
    }

    void ShowInfo(ExcelRow guest)
    {
        if (guest == null)
        {
            nameText.text = "";
            sexText.text = title1Text.text = title2Text.text = unitText.text = tableText.text = "";
        }
        else
        {
            nameText.text = guest.Name;
            sexText.text = guest.Sex;
            title1Text.text = guest.Title1;
            // tableText.text = "Bàn số: " + guest.Table;
            tableText.text = string.IsNullOrEmpty(guest.Table) ? "" : "Bàn số: " + guest.Table;
            title2Text.text = guest.Title2;
            unitText.text = guest.Unit;
        }
        StartCoroutine(ForceRefreshLayoutNextFrame());
    }

    IEnumerator ForceRefreshLayoutNextFrame()
    {
        yield return null;
        LayoutRebuilder.ForceRebuildLayoutImmediate(nameLabel.GetComponent<RectTransform>());
        LayoutRebuilder.ForceRebuildLayoutImmediate(titleLabel.GetComponent<RectTransform>());
    }

    [Serializable]
    public class ExcelRow
    {
        public int RowNumber;
        public string ID, Sex, Name, Unit, Title1, Title2, Table;
    }
}
