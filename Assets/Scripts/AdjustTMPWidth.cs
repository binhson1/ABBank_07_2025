using UnityEngine;
using TMPro;

public class AdjustTMPWidth : MonoBehaviour
{
    public TextMeshProUGUI tmpText;  // Kéo TMP vào đây
    public RectTransform rectTransform; // Kéo RectTransform vào đây hoặc lấy từ chính tmpText

    void Start()
    {
        AdjustWidth();
    }

    void Update()
    {

    }
    public void AdjustWidth()
    {
        if (tmpText == null || rectTransform == null) return;


        tmpText.ForceMeshUpdate();


        float textWidth = tmpText.preferredWidth;


        float padding = 10f;


        rectTransform.SetSizeWithCurrentAnchors(RectTransform.Axis.Horizontal, textWidth + padding);
    }
}
