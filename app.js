Office.onReady();

function scanEmailLinks() {
    var item = Office.context.mailbox.item;
    var container = document.getElementById("results");
    container.innerHTML = "<p>正在讀取電郵內容...</p>";

    // 請換成你公司嘅網址域名（例如 company.com）
    var myCompanyDomain = "urban.com.hk"; 

    // 讀取當前開啟緊嘅電郵內容
    item.body.getAsync(Office.CoercionType.Html, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            var parser = new DOMParser();
            var doc = parser.parseFromString(result.value, "text/html");
            var links = doc.querySelectorAll("a[href]");

            container.innerHTML = "";

            if (links.length === 0) {
                container.innerHTML = "<p>呢封電郵入面冇搵到任何超連結。</p>";
                return;
            }

            links.forEach(function (a) {
                var url = a.getAttribute("href").trim();

                // 只檢查 http / https 網址
                if (url.startsWith("http://") || url.startsWith("https://")) {
                    var isInternal = url.toLowerCase().includes(myCompanyDomain.toLowerCase());

                    var div = document.createElement("div");
                    div.className = "link-item " + (isInternal ? "internal" : "external");

                    var badge = isInternal 
                        ? '<span class="badge badge-safe">公司內部連結</span>' 
                        : '<span class="badge badge-danger">⚠️ 外部連結 (小心 Spam/Phishing)</span>';

                    div.innerHTML = badge + 
                        '<div><strong>文字:</strong> ' + (a.innerText || "[冇文字]") + '</div>' +
                        '<div><strong>目標網址:</strong> ' + url + '</div>' +
                        '<button class="open-btn">點擊開啟</button>';

                    // 綁定點擊事件，彈出確認對話框
                    div.querySelector(".open-btn").onclick = function () {
                        if (!isInternal) {
                            var confirmOpen = confirm(
                                "【安全警告】\n\n" +
                                "你正準備開啟一個【外部連結】！\n\n" +
                                "網址：" + url + "\n\n" +
                                "請注意：\n" +
                                "1. 切勿隨意輸入公司密碼或個人資料。\n" +
                                "2. 確認寄件者同網址係咪可信。\n\n" +
                                "確定要繼續打開嗎？"
                            );
                            if (confirmOpen) {
                                window.open(url, "_blank");
                            }
                        } else {
                            window.open(url, "_blank");
                        }
                    };

                    container.appendChild(div);
                }
            });
        } else {
            container.innerHTML = "<p style='color:red;'>讀取失敗，請重試。</p>";
        }
    });
}