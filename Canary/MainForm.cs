using System.Text.Json;
using System.Text;
using System.Diagnostics;
using System.Text.Json.Nodes;
using System.Collections.Generic;
using ClosedXML.Excel;
using System.Text.RegularExpressions;

namespace Canary
{
    public partial class MainForm : Form
    {
        private const int THREAD_COUNT = 5;

        private const string STATION_SHIBUYA = "6a7581c4-c06c-4deb-8275-688d209524c1";
        private const string STATION_AKIHABARA = "cfd3a62d-8b72-440d-899f-ddb52ad8db4b";

        private readonly string[] Prefecture =
        [
            // 東京
            "5eeb60d1-3e54-4ce1-82f9-df6848481fdf",

            // 神奈川
            "2bbacfaa-febf-491b-99cb-18c95c9684ef",

            // 千葉
            "d5758809-3ada-4c5e-a2b8-e8260a052eee",

            // 埼玉
            "d1298901-dfc5-4812-a303-534de562679e",

            // 茨城
            "3d5cc374-612e-4abb-8c0c-5105d5123ca7",

            // 栃木
            "660db406-82a0-4a6f-bb61-13a8d3fd1162",

            // 群馬
            "622ffbdc-01fe-4d57-97b3-450576003a09"
        ];

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public MainForm()
        {
            InitializeComponent();

            // 初期設定
            cmbToShibuya1.SelectedIndex = 0;
            cmbToShibuya2.SelectedIndex = 0;
            cmbToAkiba1.SelectedIndex = 0;
            cmbToAkiba2.SelectedIndex = 0;
            cmbChinryo1.SelectedIndex = 0;
            cmbChinryo2.SelectedIndex = 0;
            cmbSenyu1.SelectedIndex = 0;
            cmbSenyu2.SelectedIndex = 0;
            cmbNensu.SelectedIndex = 0;
            cmbToho.SelectedIndex = 0;

            // 各種コントロールに変更時イベントを追加する
            SetChangedEvent();
        }

        /// <summary>
        /// 各種コントロールに変更時イベントを追加する
        /// </summary>
        private void SetChangedEvent()
        {
            for (int i = 0; i < Prefecture.Length; i++)
            {
                if (Controls.Find("chkArea" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged += UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 2; i++)
            {
                if (Controls.Find("cmbToShibuya" + (i + 1).ToString(), true)[0] is ComboBox c)
                {
                    c.SelectedIndexChanged += UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 2; i++)
            {
                if (Controls.Find("cmbToAkiba" + (i + 1).ToString(), true)[0] is ComboBox c)
                {
                    c.SelectedIndexChanged += UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 2; i++)
            {
                if (Controls.Find("cmbChinryo" + (i + 1).ToString(), true)[0] is ComboBox c)
                {
                    c.SelectedIndexChanged += UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 3; i++)
            {
                if (Controls.Find("chkChinryo" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged += UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 2; i++)
            {
                if (Controls.Find("cmbSenyu" + (i + 1).ToString(), true)[0] is ComboBox c)
                {
                    c.SelectedIndexChanged += UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 1; i++)
            {
                if (Controls.Find("cmbNensu", true)[0] is ComboBox c)
                {
                    c.SelectedIndexChanged += UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 1; i++)
            {
                if (Controls.Find("cmbToho", true)[0] is ComboBox c)
                {
                    c.SelectedIndexChanged += UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 13; i++)
            {
                if (Controls.Find("chkMadori" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged += UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 3; i++)
            {
                if (Controls.Find("chkSyubetsu" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged += UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 4; i++)
            {
                if (Controls.Find("chkKouzou" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged += UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 5; i++)
            {
                if (Controls.Find("chkIchi" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged += UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 4; i++)
            {
                if (Controls.Find("chkBath" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged += UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 8; i++)
            {
                if (Controls.Find("chkHoui" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged += UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 6; i++)
            {
                if (Controls.Find("chkKitchen" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged += UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 8; i++)
            {
                if (Controls.Find("chkNyukyo" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged += UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 5; i++)
            {
                if (Controls.Find("chkSetsubi" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged += UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 8; i++)
            {
                if (Controls.Find("chkShitunai" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged += UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 5; i++)
            {
                if (Controls.Find("chkChusya" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged += UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 2; i++)
            {
                if (Controls.Find("chkReidanbou" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged += UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 5; i++)
            {
                if (Controls.Find("chkSecurity" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged += UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 2; i++)
            {
                if (Controls.Find("chkJogai" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged += UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 7; i++)
            {
                if (Controls.Find("chkSonota" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged += UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 1; i++)
            {
                if (Controls.Find("chkOriginal" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged += UpdateTaisyoKensu;
                }
            }
        }

        /// <summary>
        /// 各種コントロールに設定された変更時イベントを削除する
        /// </summary>
        private void RemoveChangedEvent()
        {
            for (int i = 0; i < Prefecture.Length; i++)
            {
                if (Controls.Find("chkArea" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged -= UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 2; i++)
            {
                if (Controls.Find("cmbToShibuya" + (i + 1).ToString(), true)[0] is ComboBox c)
                {
                    c.SelectedIndexChanged -= UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 2; i++)
            {
                if (Controls.Find("cmbToAkiba" + (i + 1).ToString(), true)[0] is ComboBox c)
                {
                    c.SelectedIndexChanged -= UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 2; i++)
            {
                if (Controls.Find("cmbChinryo" + (i + 1).ToString(), true)[0] is ComboBox c)
                {
                    c.SelectedIndexChanged -= UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 3; i++)
            {
                if (Controls.Find("chkChinryo" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged -= UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 2; i++)
            {
                if (Controls.Find("cmbSenyu" + (i + 1).ToString(), true)[0] is ComboBox c)
                {
                    c.SelectedIndexChanged -= UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 1; i++)
            {
                if (Controls.Find("cmbNensu", true)[0] is ComboBox c)
                {
                    c.SelectedIndexChanged -= UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 1; i++)
            {
                if (Controls.Find("cmbToho", true)[0] is ComboBox c)
                {
                    c.SelectedIndexChanged -= UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 13; i++)
            {
                if (Controls.Find("chkMadori" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged -= UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 3; i++)
            {
                if (Controls.Find("chkSyubetsu" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged -= UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 4; i++)
            {
                if (Controls.Find("chkKouzou" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged -= UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 5; i++)
            {
                if (Controls.Find("chkIchi" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged -= UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 4; i++)
            {
                if (Controls.Find("chkBath" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged -= UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 8; i++)
            {
                if (Controls.Find("chkHoui" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged -= UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 6; i++)
            {
                if (Controls.Find("chkKitchen" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged -= UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 8; i++)
            {
                if (Controls.Find("chkNyukyo" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged -= UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 5; i++)
            {
                if (Controls.Find("chkSetsubi" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged -= UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 8; i++)
            {
                if (Controls.Find("chkShitunai" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged -= UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 5; i++)
            {
                if (Controls.Find("chkChusya" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged -= UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 2; i++)
            {
                if (Controls.Find("chkReidanbou" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged -= UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 5; i++)
            {
                if (Controls.Find("chkSecurity" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged -= UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 2; i++)
            {
                if (Controls.Find("chkJogai" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged -= UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 7; i++)
            {
                if (Controls.Find("chkSonota" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged -= UpdateTaisyoKensu;
                }
            }

            for (int i = 0; i < 1; i++)
            {
                if (Controls.Find("chkOriginal" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.CheckedChanged -= UpdateTaisyoKensu;
                }
            }
        }

        /// <summary>
        /// 対象件数を取得
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void UpdateTaisyoKensu(object? sender, EventArgs e)
        {
            Task.Run(() =>
            {
                bool isExit = false;

                Invoke(() =>
                {
                    // エリアがどこも選択されていない場合は終了
                    if (chkArea1.Checked == false && chkArea2.Checked == false && chkArea3.Checked == false &&
                        chkArea4.Checked == false && chkArea5.Checked == false && chkArea6.Checked == false && chkArea7.Checked == false)
                    {
                        isExit = true;
                    }
                });

                if (isExit == true)
                {
                    Invoke(() =>
                    {
                        // 対象件数を取得
                        edtKensu.Text = "0件";
                    });

                    return;
                }

                // 時間計測用
                var sw = new Stopwatch();

                // 計測開始
                sw.Start();

                List<JsonNode> result_estates = [];
                List<JsonNode> result_rooms = [];
                List<IEnumerable<int>> chunks = [];
                Parameter2 param = null;

                Invoke(() =>
                {
                    // 検索条件を設定する
                    List<string> layoutNames = [];
                    for (int i = 0; i < 13; i++)
                    {
                        CheckBox? cs = Controls.Find("chkMadori" + (i + 1).ToString(), true)[0] as CheckBox;
                        if (cs?.Checked == true)
                        {
                            layoutNames.Add(cs.Text);
                        }
                    }

                    List<string> prefectureIds = [];
                    for (int i = 0; i < Prefecture.Length; i++)
                    {
                        CheckBox? cs = Controls.Find("chkArea" + (i + 1).ToString(), true)[0] as CheckBox;
                        if (cs?.Checked == true)
                        {
                            prefectureIds.Add(Prefecture[i]);
                        }
                    }

                    // 検索条件識別子リストを取得する
                    Dictionary<string, string> id_list = [];
                    foreach (var elem in File.ReadAllLines("id_list.csv"))
                    {
                        // カンマで分割する
                        var split_data = elem.Split(",");

                        // 辞書に登録する
                        id_list.Add(split_data[0], split_data[1]);
                    }

                    List<string> searchOptions = [];

                    for (int i = 0; i < 2; i++)
                    {
                        if (Controls.Find("chkChinryo" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                        {
                            searchOptions.Add(id_list[c.Text]);
                        }
                    }

                    for (int i = 0; i < 3; i++)
                    {
                        if (Controls.Find("chkSyubetsu" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                        {
                            searchOptions.Add(id_list[c.Text]);
                        }
                    }

                    for (int i = 0; i < 4; i++)
                    {
                        if (Controls.Find("chkKouzou" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                        {
                            searchOptions.Add(id_list[c.Text]);
                        }
                    }

                    for (int i = 0; i < 5; i++)
                    {
                        if (Controls.Find("chkIchi" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                        {
                            searchOptions.Add(id_list[c.Text]);
                        }
                    }

                    for (int i = 0; i < 4; i++)
                    {
                        if (Controls.Find("chkBath" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                        {
                            searchOptions.Add(id_list[c.Text]);
                        }
                    }

                    for (int i = 0; i < 8; i++)
                    {
                        if (Controls.Find("chkHoui" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                        {
                            searchOptions.Add(id_list[c.Text]);
                        }
                    }

                    for (int i = 0; i < 6; i++)
                    {
                        if (Controls.Find("chkKitchen" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                        {
                            searchOptions.Add(id_list[c.Text]);
                        }
                    }

                    for (int i = 0; i < 8; i++)
                    {
                        if (Controls.Find("chkNyukyo" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                        {
                            searchOptions.Add(id_list[c.Text]);
                        }
                    }

                    for (int i = 0; i < 5; i++)
                    {
                        if (Controls.Find("chkSetsubi" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                        {
                            searchOptions.Add(id_list[c.Text]);
                        }
                    }

                    for (int i = 0; i < 8; i++)
                    {
                        if (Controls.Find("chkShitunai" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                        {
                            searchOptions.Add(id_list[c.Text]);
                        }
                    }

                    for (int i = 0; i < 5; i++)
                    {
                        if (Controls.Find("chkChusya" + (i + 1).ToString(), true)[0] is CheckBox c)
                            if (c != null && c.Checked)
                            {
                                searchOptions.Add(id_list[c.Text]);
                            }
                    }

                    for (int i = 0; i < 2; i++)
                    {
                        if (Controls.Find("chkReidanbou" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                        {
                            searchOptions.Add(id_list[c.Text]);
                        }
                    }

                    for (int i = 0; i < 5; i++)
                    {
                        if (Controls.Find("chkSecurity" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                        {
                            searchOptions.Add(id_list[c.Text]);
                        }
                    }

                    for (int i = 0; i < 2; i++)
                    {
                        if (Controls.Find("chkJogai" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                        {
                            searchOptions.Add(id_list[c.Text]);
                        }
                    }

                    for (int i = 0; i < 7; i++)
                    {
                        if (Controls.Find("chkSonota" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                        {
                            searchOptions.Add(id_list[c.Text]);
                        }
                    }

                    // 対象件数を取得する
                    var taisyo_kensu = SearchTaisyoKensu();

                    if (taisyo_kensu <= 0)
                    {
                        return;
                    }

                    // 対象を分割する

                    var split_count = taisyo_kensu / 20 + 1;

                    List<int> tmp = [];
                    for (int i = 0; i < taisyo_kensu; i += 20)
                    {
                        tmp.Add(i);
                    }

                    chunks = tmp.Select((v, i) => new { v, i })
                        .GroupBy(x => x.i / (1 + tmp.Count / THREAD_COUNT))
                        .Select(g => g.Select(x => x.v)).ToList();

                    // パラメータを取得
                    param = new()
                    {
                        chomeiIds = Array.Empty<string>(),
                        cityIds = Array.Empty<string>(),
                        commutes =
                        [
                            new()
                            {
                                sourceStationId = STATION_SHIBUYA,
                                timeMinutes = int.Parse(cmbToShibuya1.Text.Replace("分", "")),
                                changeCount = new Changecount()
                                {
                                    value = cmbToShibuya2.SelectedIndex == 0 ? -1 : int.Parse(cmbToShibuya2.Text.Replace("回", "")),
                                    hasValue = cmbToShibuya2.SelectedIndex != 0
                                }
                            },
                            new()
                            {
                                sourceStationId = STATION_AKIHABARA,
                                timeMinutes = int.Parse(cmbToAkiba1.Text.Replace("分", "")),
                                changeCount = new Changecount()
                                {
                                    value = cmbToAkiba2.SelectedIndex == 0 ? -1 : int.Parse(cmbToAkiba2.Text.Replace("回", "")),
                                    hasValue = cmbToAkiba2.SelectedIndex != 0
                                }
                            }
                        ],
                        duringMax = new Duringmax()
                        {
                            hasValue = cmbToho.SelectedIndex != 0,
                            value = cmbToho.SelectedIndex == 0 ? 0 : int.Parse(cmbToho.Text.Replace("分以下", ""))
                        },
                        includeAdminFee = chkChinryo3.Checked,
                        isNewArrival = false,
                        keywords = Array.Empty<string>(),
                        layoutNames = [.. layoutNames],
                        oldMax = new Oldmax()
                        {
                            hasValue = cmbNensu.SelectedIndex != 0,
                            value = cmbNensu.SelectedIndex == 0 ? 0 : int.Parse(cmbNensu.Text.Replace("築", "").Replace("年以下", ""))
                        },
                        prefectureIds = [.. prefectureIds],
                        rentMax = new Rentmax()
                        {
                            hasValue = cmbChinryo2.SelectedIndex != 0,
                            value = cmbChinryo2.SelectedIndex == 0 ? 0 : int.Parse(cmbChinryo2.Text.Replace("万以下", "")) * 10000
                        },
                        rentMin = new Rentmin()
                        {
                            hasValue = cmbChinryo1.SelectedIndex != 0,
                            value = cmbChinryo1.SelectedIndex == 0 ? 0 : int.Parse(cmbChinryo1.Text.Replace("万以上", "")) * 10000
                        },
                        searchOptionIds = [.. searchOptions],
                        shatakuCode = new Shatakucode()
                        {
                            hasValue = false,
                            value = ""
                        },
                        squareMax = new Squaremax()
                        {
                            hasValue = cmbSenyu2.SelectedIndex != 0,
                            value = cmbSenyu2.SelectedIndex == 0 ? 0 : int.Parse(cmbSenyu2.Text.Replace("㎡以下", ""))
                        },
                        squareMin = new Squaremin()
                        {
                            hasValue = cmbSenyu1.SelectedIndex != 0,
                            value = cmbSenyu1.SelectedIndex == 0 ? 0 : int.Parse(cmbSenyu1.Text.Replace("㎡以上", ""))
                        },
                        stationIds = Array.Empty<string>(),
                        limit = 20,
                        listType = 2,
                        offset = new Offset()
                        {
                            hasValue = true,
                            value = "0"
                        },
                        searchPurpose = true,
                        searchSessionId = string.Empty,
                        sortType = 1
                    };

                });

                if (chunks.Count <= 0)
                {
                    Invoke(() =>
                    {
                        // 対象件数を取得
                        edtKensu.Text = "0件";
                    });

                    return;
                }

                // 対象の部屋が存在しなくなるまでループ
                List<Task<Tuple<List<JsonNode>, List<JsonNode>>>> task_list = [];

                for (int i = 0; i < chunks.Count; i++)
                {
                    // オフセットリスト
                    var offset_list = chunks[i];

                    // パラメータをコピー
                    Parameter2 param_tmp = new(param);

                    var index = i;

                    task_list.Add(Task.Run(() =>
                    {
                        List<JsonNode> result1 = [];
                        List<JsonNode> result2 = [];

                        for (int j = 0; j < offset_list.Count(); j++)
                        {
                            Debug.WriteLine("スレッド番号 = " + (index + 1).ToString() + " : オフセット = " + offset_list.ToArray()[j].ToString());

                            // 個別のパラメータを指定
                            param_tmp.offset = new Offset()
                            {
                                hasValue = true,
                                value = offset_list.ToArray()[j].ToString()
                            };

                            using var client = new HttpClient();

                            // APIを利用して検索する
                            var task = client.PostAsync("https://api.user.canary-app.com/v1/chintaiRooms:search", new StringContent(JsonSerializer.Serialize(param_tmp), Encoding.UTF8, "application/json"));
                            task.Wait();

                            // HttpContentから文字列を取得する
                            var task2 = task.Result.Content.ReadAsStringAsync();
                            task2.Wait();

                            // Jsonオブジェクトにパースする
                            var json_data = JsonNode.Parse(task2.Result);

                            // 配列長を取得
                            var length1 = json_data["chintaiEstates"].AsArray().Count;
                            var length2 = json_data["chintaiRooms"].AsArray().Count;

                            if (length1 <= 0 || length2 <= 0)
                            {
                                Debug.WriteLine("スレッド番号 = " + (index + 1).ToString() + " : ループをBREAK");
                                break;
                            }

                            for (int k = 0; k < length1; k++)
                            {
                                result1.Add(json_data["chintaiEstates"][k]);
                            }

                            for (int k = 0; k < length2; k++)
                            {
                                result2.Add(json_data["chintaiRooms"][k]);
                            }
                        }

                        return new Tuple<List<JsonNode>, List<JsonNode>>(result1, result2);
                    }));
                }

                // すべてのスレッドが完了するまで待機する
                for (int i = 0; i < task_list.Count; i++)
                {
                    task_list[i].Wait();
                    result_estates.AddRange(task_list[i].Result.Item1);
                    result_rooms.AddRange(task_list[i].Result.Item2);
                }

                // 計測停止
                sw.Stop();

                TimeSpan ts = sw.Elapsed;
                Debug.WriteLine($"　{ts.Hours}時間 {ts.Minutes}分 {ts.Seconds}秒 {ts.Milliseconds}ミリ秒");
                Debug.WriteLine(result_estates.Count.ToString() + "件");

                Invoke(() =>
                {
                    // オリジナル条件を反映する
                    List<JsonNode> extract_rooms = new();
                    for (int i = 0; i < result_rooms.Count; i++)
                    {
                        if (chkOriginal1.Checked && result_rooms[i]["layoutDetail"]["value"].ToString().Contains("和"))
                        {
                            continue;
                        }

                        extract_rooms.Add(result_rooms[i]);
                    }

                    // 対象件数を取得
                    edtKensu.Text = extract_rooms.Count.ToString() + "件";
                });
            });
        }

        /// <summary>
        /// 検索ボタン押下時の処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOutput_Click(object sender, EventArgs e)
        {
            // 賃貸リストを生成する
            var result = GenerateChintaiList();

            if (result == null)
            {
                return;
            }

            // エクセルに落とし込む
            SaveChintaiData(result.Item1, result.Item2);
        }

        /// <summary>
        /// 賃貸情報をエクセルファイルに保存する
        /// </summary>
        /// <param name="chintai_data"></param>
        private void SaveChintaiData(List<JsonNode> estates_data, List<JsonNode> rooms_data, bool isPreview = false)
        {
            // エクセルを生成
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("賃貸情報");

            // ヘッダー生成
            var header = new[] { "No.", "建物名", "都道府県", "住所", "建築年", "総階数(地下含む)", "建物種類", "構造", "部屋番号", "階数",
                "敷金", "礼金", "家賃(共益費込)", "入居可能タイミング", "レイアウト", "面積", "更新日時", "間取り", "ページURL", "間取り画像"};
            ws.Cell("A1").InsertData(new[] { header });

            for (int i = 0; i < rooms_data.Count; i++)
            {
                // 建物情報を取得する
                JsonNode estates = null;
                for (int j = 0; j < estates_data.Count; j++)
                {
                    if (rooms_data[i]["chintaiEstateId"].ToString() == estates_data[j]["id"].ToString())
                    {
                        estates = estates_data[j];
                        break;
                    }
                }

                if (estates == null)
                {
                    continue;
                }

                // 住所を区切る
                var reg = new Regex("(...??[都道府県])((?:旭川|伊達|石狩|盛岡|奥州|田村|南相馬|那須塩原|東村山|武蔵村山|羽村|十日町|上越|富山|野々市|大町|蒲郡|四日市|姫路|大和郡山|廿日市|下松|岩国|田川|大村)市|.+?郡(?:玉村|大町|.+?)[町村]|.+?市.+?区|.+?[市区町村])(.+)");
                var s = reg.Split(estates["addressStr"].ToString())[1];


                var row = new List<string>
                {
                    (i + 1).ToString(),
                    estates["name"].ToString(),
                    s,
                    estates["addressStr"].ToString(),
                    estates["builtAtYear"].ToString(),
                    (int.Parse(estates["aboveGroundStory"].ToString()) + int.Parse(estates["underGroundStory"].ToString())).ToString(),
                    estates["estateTypeName"].ToString(),
                    estates["structureName"].ToString(),
                    rooms_data[i]["roomNumber"].ToString(),
                    rooms_data[i]["floor"].ToString(),
                    rooms_data[i]["securityDeposit"].ToString(),
                    rooms_data[i]["keyMoney"].ToString(),
                    (int.Parse(rooms_data[i]["rent"].ToString()) + int.Parse(rooms_data[i]["adminFee"].ToString())).ToString(),
                    rooms_data[i]["moveIn"].ToString(),
                    rooms_data[i]["layout"]["name"].ToString(),
                    rooms_data[i]["square"].ToString(),
                    rooms_data[i]["updatedAt"].ToString(),
                    rooms_data[i]["layoutDetail"]["value"].ToString()
                };

                // 行挿入
                ws.Cell("A" + (i + 2).ToString()).InsertData(new[] { row });

                // ハイパーリンク挿入
                var url = "https://web.canary-app.jp/chintai/rooms/" + rooms_data[i]["id"].ToString();
                ws.Cell(i + 2, row.Count + 1).SetHyperlink(new XLHyperlink(url));
                ws.Cell(i + 2, row.Count + 1).Value = url;
                try
                {
                    ws.Cell(i + 2, row.Count + 2).SetHyperlink(new XLHyperlink(rooms_data[i]["layoutImageUrl"]["value"].ToString()));
                    ws.Cell(i + 2, row.Count + 2).Value = rooms_data[i]["layoutImageUrl"]["value"].ToString();
                }
                catch
                {
                    ws.Cell(i + 2, row.Count + 2).Value = rooms_data[i]["layoutImageUrl"]["value"].ToString();
                }
            }

            // 列を整形する
            ws.ColumnsUsed().AdjustToContents();

            if (isPreview == false)
            {
                // 保存する
                SaveFileDialog sfd = new()
                {
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                    Filter = "xlsxファイル(*.xlsx)|*.xlsx",
                    Title = "保存先のファイルを選択してください"
                };

                //ダイアログを表示する
                if (sfd.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                // ファイルを保存する
                wb.SaveAs(sfd.FileName);
            }
            else
            {
                // ファイルを保存する
                wb.SaveAs("tmp.xlsx");

                // 画面に表示する
                ExcelViewer viewer = new ExcelViewer("tmp.xlsx", header.Length);
                viewer.ShowDialog();

                if (File.Exists("tmp.xlsx"))
                {
                    File.Delete("tmp.xlsx");
                }
            }
        }

        /// <summary>
        /// 指定した条件に合致する物件数を取得する
        /// </summary>
        /// <returns></returns>
        private int SearchTaisyoKensu()
        {
            Parameter1 param = new();

            Invoke(() =>
            {
                // エリアがどこも選択されていない場合は終了
                if (chkArea1.Checked == true || chkArea2.Checked == true || chkArea3.Checked == true ||
                    chkArea4.Checked == true || chkArea5.Checked == true || chkArea6.Checked == true || chkArea7.Checked == true)
                {

                    // 検索条件を設定する
                    List<string> layoutNames = [];
                    for (int i = 0; i < 13; i++)
                    {
                        CheckBox? cs = Controls.Find("chkMadori" + (i + 1).ToString(), true)[0] as CheckBox;
                        if (cs?.Checked == true)
                        {
                            layoutNames.Add(cs.Text);
                        }
                    }

                    List<string> prefectureIds = [];
                    for (int i = 0; i < Prefecture.Length; i++)
                    {
                        CheckBox? cs = Controls.Find("chkArea" + (i + 1).ToString(), true)[0] as CheckBox;
                        if (cs?.Checked == true)
                        {
                            prefectureIds.Add(Prefecture[i]);
                        }
                    }

                    // 検索条件識別子リストを取得する
                    Dictionary<string, string> id_list = [];
                    foreach (var elem in File.ReadAllLines("id_list.csv"))
                    {
                        // カンマで分割する
                        var split_data = elem.Split(",");

                        // 辞書に登録する
                        id_list.Add(split_data[0], split_data[1]);
                    }

                    List<string> searchOptions = [];

                    for (int i = 0; i < 2; i++)
                    {
                        if (Controls.Find("chkChinryo" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                        {
                            searchOptions.Add(id_list[c.Text]);
                        }
                    }

                    for (int i = 0; i < 3; i++)
                    {
                        if (Controls.Find("chkSyubetsu" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                        {
                            searchOptions.Add(id_list[c.Text]);
                        }
                    }

                    for (int i = 0; i < 4; i++)
                    {
                        if (Controls.Find("chkKouzou" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                        {
                            searchOptions.Add(id_list[c.Text]);
                        }
                    }

                    for (int i = 0; i < 5; i++)
                    {
                        if (Controls.Find("chkIchi" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                        {
                            searchOptions.Add(id_list[c.Text]);
                        }
                    }

                    for (int i = 0; i < 4; i++)
                    {
                        if (Controls.Find("chkBath" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                        {
                            searchOptions.Add(id_list[c.Text]);
                        }
                    }

                    for (int i = 0; i < 8; i++)
                    {
                        if (Controls.Find("chkHoui" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                        {
                            searchOptions.Add(id_list[c.Text]);
                        }
                    }

                    for (int i = 0; i < 6; i++)
                    {
                        if (Controls.Find("chkKitchen" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                        {
                            searchOptions.Add(id_list[c.Text]);
                        }
                    }

                    for (int i = 0; i < 8; i++)
                    {
                        if (Controls.Find("chkNyukyo" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                        {
                            searchOptions.Add(id_list[c.Text]);
                        }
                    }

                    for (int i = 0; i < 5; i++)
                    {
                        if (Controls.Find("chkSetsubi" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                        {
                            searchOptions.Add(id_list[c.Text]);
                        }
                    }

                    for (int i = 0; i < 8; i++)
                    {
                        if (Controls.Find("chkShitunai" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                        {
                            searchOptions.Add(id_list[c.Text]);
                        }
                    }

                    for (int i = 0; i < 5; i++)
                    {
                        if (Controls.Find("chkChusya" + (i + 1).ToString(), true)[0] is CheckBox c)
                            if (c != null && c.Checked)
                            {
                                searchOptions.Add(id_list[c.Text]);
                            }
                    }

                    for (int i = 0; i < 2; i++)
                    {
                        if (Controls.Find("chkReidanbou" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                        {
                            searchOptions.Add(id_list[c.Text]);
                        }
                    }

                    for (int i = 0; i < 5; i++)
                    {
                        if (Controls.Find("chkSecurity" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                        {
                            searchOptions.Add(id_list[c.Text]);
                        }
                    }

                    for (int i = 0; i < 2; i++)
                    {
                        if (Controls.Find("chkJogai" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                        {
                            searchOptions.Add(id_list[c.Text]);
                        }
                    }

                    for (int i = 0; i < 7; i++)
                    {
                        if (Controls.Find("chkSonota" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                        {
                            searchOptions.Add(id_list[c.Text]);
                        }
                    }

                    param = new Parameter1()
                    {
                        chomeiIds = Array.Empty<string>(),
                        cityIds = Array.Empty<string>(),
                        commutes =
                        [
                            new()
                            {
                                sourceStationId = STATION_SHIBUYA,
                                timeMinutes = int.Parse(cmbToShibuya1.Text.Replace("分", "")),
                                changeCount = new Changecount()
                                {
                                    value = cmbToShibuya2.SelectedIndex == 0 ? -1 : int.Parse(cmbToShibuya2.Text.Replace("回", "")),
                                    hasValue = cmbToShibuya2.SelectedIndex != 0
                                }
                            },
                            new()
                            {
                                sourceStationId = STATION_AKIHABARA,
                                timeMinutes = int.Parse(cmbToAkiba1.Text.Replace("分", "")),
                                changeCount = new Changecount()
                                {
                                    value = cmbToAkiba2.SelectedIndex == 0 ? -1 : int.Parse(cmbToAkiba2.Text.Replace("回", "")),
                                    hasValue = cmbToAkiba2.SelectedIndex != 0
                                }
                            }
                        ],
                        duringMax = new Duringmax()
                        {
                            hasValue = cmbToho.SelectedIndex != 0,
                            value = cmbToho.SelectedIndex == 0 ? 0 : int.Parse(cmbToho.Text.Replace("分以下", ""))
                        },
                        includeAdminFee = chkChinryo3.Checked,
                        isNewArrival = false,
                        keywords = Array.Empty<string>(),
                        layoutNames = [.. layoutNames],
                        oldMax = new Oldmax()
                        {
                            hasValue = cmbNensu.SelectedIndex != 0,
                            value = cmbNensu.SelectedIndex == 0 ? 0 : int.Parse(cmbNensu.Text.Replace("築", "").Replace("年以下", ""))
                        },
                        prefectureIds = [.. prefectureIds],
                        rentMax = new Rentmax()
                        {
                            hasValue = cmbChinryo2.SelectedIndex != 0,
                            value = cmbChinryo2.SelectedIndex == 0 ? 0 : int.Parse(cmbChinryo2.Text.Replace("万以下", "")) * 10000
                        },
                        rentMin = new Rentmin()
                        {
                            hasValue = cmbChinryo1.SelectedIndex != 0,
                            value = cmbChinryo1.SelectedIndex == 0 ? 0 : int.Parse(cmbChinryo1.Text.Replace("万以上", "")) * 10000
                        },
                        searchOptionIds = [.. searchOptions],
                        shatakuCode = new Shatakucode()
                        {
                            hasValue = false,
                            value = ""
                        },
                        squareMax = new Squaremax()
                        {
                            hasValue = cmbSenyu2.SelectedIndex != 0,
                            value = cmbSenyu2.SelectedIndex == 0 ? 0 : int.Parse(cmbSenyu2.Text.Replace("㎡以下", ""))
                        },
                        squareMin = new Squaremin()
                        {
                            hasValue = cmbSenyu1.SelectedIndex != 0,
                            value = cmbSenyu1.SelectedIndex == 0 ? 0 : int.Parse(cmbSenyu1.Text.Replace("㎡以上", ""))
                        },
                        stationIds = Array.Empty<string>()
                    };
                }
            });

            if (param == null)
            {
                return 0;
            }

            using var client = new HttpClient();

            // APIを利用して対象件数を取得する
            var task = client.PostAsync("https://api.user.canary-app.com/v1/chintaiRooms:countBySearchFilter", new StringContent(JsonSerializer.Serialize(param), Encoding.UTF8, "application/json"));
            task.Wait();

            // HttpContentから文字列を取得する
            var task2 = task.Result.Content.ReadAsStringAsync();
            task2.Wait();

            // 対象件数を取り出す
            return JsonSerializer.Deserialize<TotalCount>(task2.Result).totalCount;
        }

        /// <summary>
        /// コントロールの状態を保存する
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSave_Click(object sender, EventArgs e)
        {
            Dictionary<string, string> controlData = [];

            for (int i = 0; i < Prefecture.Length; i++)
            {
                if (Controls.Find("chkArea" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    controlData.Add(c.Name, c.Checked.ToString());
                }
            }

            for (int i = 0; i < 2; i++)
            {
                if (Controls.Find("cmbToShibuya" + (i + 1).ToString(), true)[0] is ComboBox c)
                {
                    controlData.Add(c.Name, c.SelectedIndex.ToString());
                }
            }

            for (int i = 0; i < 2; i++)
            {
                if (Controls.Find("cmbToAkiba" + (i + 1).ToString(), true)[0] is ComboBox c)
                {
                    controlData.Add(c.Name, c.SelectedIndex.ToString());
                }
            }

            for (int i = 0; i < 2; i++)
            {
                if (Controls.Find("cmbChinryo" + (i + 1).ToString(), true)[0] is ComboBox c)
                {
                    controlData.Add(c.Name, c.SelectedIndex.ToString());
                }
            }

            for (int i = 0; i < 3; i++)
            {
                if (Controls.Find("chkChinryo" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    controlData.Add(c.Name, c.Checked.ToString());
                }
            }

            for (int i = 0; i < 2; i++)
            {
                if (Controls.Find("cmbSenyu" + (i + 1).ToString(), true)[0] is ComboBox c)
                {
                    controlData.Add(c.Name, c.SelectedIndex.ToString());
                }
            }

            for (int i = 0; i < 1; i++)
            {
                if (Controls.Find("cmbNensu", true)[0] is ComboBox c)
                {
                    controlData.Add(c.Name, c.SelectedIndex.ToString());
                }
            }

            for (int i = 0; i < 1; i++)
            {
                if (Controls.Find("cmbToho", true)[0] is ComboBox c)
                {
                    controlData.Add(c.Name, c.SelectedIndex.ToString());
                }
            }

            for (int i = 0; i < 13; i++)
            {
                if (Controls.Find("chkMadori" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    controlData.Add(c.Name, c.Checked.ToString());
                }
            }

            for (int i = 0; i < 3; i++)
            {
                if (Controls.Find("chkSyubetsu" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    controlData.Add(c.Name, c.Checked.ToString());
                }
            }

            for (int i = 0; i < 4; i++)
            {
                if (Controls.Find("chkKouzou" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    controlData.Add(c.Name, c.Checked.ToString());
                }
            }

            for (int i = 0; i < 5; i++)
            {
                if (Controls.Find("chkIchi" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    controlData.Add(c.Name, c.Checked.ToString());
                }
            }

            for (int i = 0; i < 4; i++)
            {
                if (Controls.Find("chkBath" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    controlData.Add(c.Name, c.Checked.ToString());
                }
            }

            for (int i = 0; i < 8; i++)
            {
                if (Controls.Find("chkHoui" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    controlData.Add(c.Name, c.Checked.ToString());
                }
            }

            for (int i = 0; i < 6; i++)
            {
                if (Controls.Find("chkKitchen" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    controlData.Add(c.Name, c.Checked.ToString());
                }
            }

            for (int i = 0; i < 8; i++)
            {
                if (Controls.Find("chkNyukyo" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    controlData.Add(c.Name, c.Checked.ToString());
                }
            }

            for (int i = 0; i < 5; i++)
            {
                if (Controls.Find("chkSetsubi" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    controlData.Add(c.Name, c.Checked.ToString());
                }
            }

            for (int i = 0; i < 8; i++)
            {
                if (Controls.Find("chkShitunai" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    controlData.Add(c.Name, c.Checked.ToString());
                }
            }

            for (int i = 0; i < 5; i++)
            {
                if (Controls.Find("chkChusya" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    controlData.Add(c.Name, c.Checked.ToString());
                }
            }

            for (int i = 0; i < 2; i++)
            {
                if (Controls.Find("chkReidanbou" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    controlData.Add(c.Name, c.Checked.ToString());
                }
            }

            for (int i = 0; i < 5; i++)
            {
                if (Controls.Find("chkSecurity" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    controlData.Add(c.Name, c.Checked.ToString());
                }
            }

            for (int i = 0; i < 2; i++)
            {
                if (Controls.Find("chkJogai" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    controlData.Add(c.Name, c.Checked.ToString());
                }
            }

            for (int i = 0; i < 7; i++)
            {
                if (Controls.Find("chkSonota" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    controlData.Add(c.Name, c.Checked.ToString());
                }
            }

            for (int i = 0; i < 1; i++)
            {
                if (Controls.Find("chkOriginal" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    controlData.Add(c.Name, c.Checked.ToString());
                }
            }

            // データをシリアライズする
            var data = JsonSerializer.Serialize(controlData);

            // 保存する
            SaveFileDialog sfd = new()
            {
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                Filter = "savファイル(*.sav)|*.sav",
                Title = "保存先のファイルを選択してください"
            };

            //ダイアログを表示する
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                File.WriteAllText(sfd.FileName, data);
            }
        }

        /// <summary>
        /// コントロールの状態を復元する
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnLoad_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new()
            {
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                Filter = "savファイル(*.sav)|*.sav",
                Title = "開くファイルを選択してください"
            };

            //ダイアログを表示する
            if (ofd.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            // 一度変更イベントを削除する
            RemoveChangedEvent();

            // 指定したファイルをデシリアライズする
            var controlData = JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(ofd.FileName));

            if (controlData == null)
            {
                return;
            }

            for (int i = 0; i < Prefecture.Length; i++)
            {
                if (Controls.Find("chkArea" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.Checked = bool.Parse(controlData[c.Name]);
                }
            }

            for (int i = 0; i < 2; i++)
            {
                if (Controls.Find("cmbToShibuya" + (i + 1).ToString(), true)[0] is ComboBox c)
                {
                    c.SelectedIndex = int.Parse(controlData[c.Name]);
                }
            }

            for (int i = 0; i < 2; i++)
            {
                if (Controls.Find("cmbToAkiba" + (i + 1).ToString(), true)[0] is ComboBox c)
                {
                    c.SelectedIndex = int.Parse(controlData[c.Name]);
                }
            }

            for (int i = 0; i < 2; i++)
            {
                if (Controls.Find("cmbChinryo" + (i + 1).ToString(), true)[0] is ComboBox c)
                {
                    c.SelectedIndex = int.Parse(controlData[c.Name]);
                }
            }

            for (int i = 0; i < 3; i++)
            {
                if (Controls.Find("chkChinryo" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.Checked = bool.Parse(controlData[c.Name]);
                }
            }

            for (int i = 0; i < 2; i++)
            {
                if (Controls.Find("cmbSenyu" + (i + 1).ToString(), true)[0] is ComboBox c)
                {
                    c.SelectedIndex = int.Parse(controlData[c.Name]);
                }
            }

            for (int i = 0; i < 1; i++)
            {
                if (Controls.Find("cmbNensu", true)[0] is ComboBox c)
                {
                    c.SelectedIndex = int.Parse(controlData[c.Name]);
                }
            }

            for (int i = 0; i < 1; i++)
            {
                if (Controls.Find("cmbToho", true)[0] is ComboBox c)
                {
                    c.SelectedIndex = int.Parse(controlData[c.Name]);
                }
            }

            for (int i = 0; i < 13; i++)
            {
                if (Controls.Find("chkMadori" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.Checked = bool.Parse(controlData[c.Name]);
                }
            }

            for (int i = 0; i < 3; i++)
            {
                if (Controls.Find("chkSyubetsu" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.Checked = bool.Parse(controlData[c.Name]);
                }
            }

            for (int i = 0; i < 4; i++)
            {
                if (Controls.Find("chkKouzou" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.Checked = bool.Parse(controlData[c.Name]);
                }
            }

            for (int i = 0; i < 5; i++)
            {
                if (Controls.Find("chkIchi" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.Checked = bool.Parse(controlData[c.Name]);
                }
            }

            for (int i = 0; i < 4; i++)
            {
                if (Controls.Find("chkBath" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.Checked = bool.Parse(controlData[c.Name]);
                }
            }

            for (int i = 0; i < 8; i++)
            {
                if (Controls.Find("chkHoui" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.Checked = bool.Parse(controlData[c.Name]);
                }
            }

            for (int i = 0; i < 6; i++)
            {
                if (Controls.Find("chkKitchen" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.Checked = bool.Parse(controlData[c.Name]);
                }
            }

            for (int i = 0; i < 8; i++)
            {
                if (Controls.Find("chkNyukyo" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.Checked = bool.Parse(controlData[c.Name]);
                }
            }

            for (int i = 0; i < 5; i++)
            {
                if (Controls.Find("chkSetsubi" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.Checked = bool.Parse(controlData[c.Name]);
                }
            }

            for (int i = 0; i < 8; i++)
            {
                if (Controls.Find("chkShitunai" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.Checked = bool.Parse(controlData[c.Name]);
                }
            }

            for (int i = 0; i < 5; i++)
            {
                if (Controls.Find("chkChusya" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.Checked = bool.Parse(controlData[c.Name]);
                }
            }

            for (int i = 0; i < 2; i++)
            {
                if (Controls.Find("chkReidanbou" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.Checked = bool.Parse(controlData[c.Name]);
                }
            }

            for (int i = 0; i < 5; i++)
            {
                if (Controls.Find("chkSecurity" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.Checked = bool.Parse(controlData[c.Name]);
                }
            }

            for (int i = 0; i < 2; i++)
            {
                if (Controls.Find("chkJogai" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.Checked = bool.Parse(controlData[c.Name]);
                }
            }

            for (int i = 0; i < 7; i++)
            {
                if (Controls.Find("chkSonota" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.Checked = bool.Parse(controlData[c.Name]);
                }
            }

            for (int i = 0; i < 1; i++)
            {
                if (Controls.Find("chkOriginal" + (i + 1).ToString(), true)[0] is CheckBox c)
                {
                    c.Checked = bool.Parse(controlData[c.Name]);
                }
            }

            // 変更イベントを再度追加する
            SetChangedEvent();

            // 検索対象件数を更新する
            UpdateTaisyoKensu(sender, e);
        }

        /// <summary>
        /// 賃貸リストを生成する
        /// </summary>
        /// <returns></returns>
        private Tuple<List<JsonNode>, List<JsonNode>> GenerateChintaiList()
        {
            bool isExit = false;

            Invoke(() =>
            {
                // エリアがどこも選択されていない場合は終了
                if (chkArea1.Checked == false && chkArea2.Checked == false && chkArea3.Checked == false &&
                    chkArea4.Checked == false && chkArea5.Checked == false && chkArea6.Checked == false && chkArea7.Checked == false)
                {
                    isExit = true;
                }
            });

            if (isExit == true)
            {
                return null;
            }

            // 時間計測用
            var sw = new Stopwatch();

            // 計測開始
            sw.Start();

            List<JsonNode> result_estates = [];
            List<JsonNode> result_rooms = [];
            List<IEnumerable<int>> chunks = [];
            Parameter2 param = null;

            Invoke(() =>
            {
                // 検索条件を設定する
                List<string> layoutNames = [];
                for (int i = 0; i < 13; i++)
                {
                    CheckBox? cs = Controls.Find("chkMadori" + (i + 1).ToString(), true)[0] as CheckBox;
                    if (cs?.Checked == true)
                    {
                        layoutNames.Add(cs.Text);
                    }
                }

                List<string> prefectureIds = [];
                for (int i = 0; i < Prefecture.Length; i++)
                {
                    CheckBox? cs = Controls.Find("chkArea" + (i + 1).ToString(), true)[0] as CheckBox;
                    if (cs?.Checked == true)
                    {
                        prefectureIds.Add(Prefecture[i]);
                    }
                }

                // 検索条件識別子リストを取得する
                Dictionary<string, string> id_list = [];
                foreach (var elem in File.ReadAllLines("id_list.csv"))
                {
                    // カンマで分割する
                    var split_data = elem.Split(",");

                    // 辞書に登録する
                    id_list.Add(split_data[0], split_data[1]);
                }

                List<string> searchOptions = [];

                for (int i = 0; i < 2; i++)
                {
                    if (Controls.Find("chkChinryo" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                    {
                        searchOptions.Add(id_list[c.Text]);
                    }
                }

                for (int i = 0; i < 3; i++)
                {
                    if (Controls.Find("chkSyubetsu" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                    {
                        searchOptions.Add(id_list[c.Text]);
                    }
                }

                for (int i = 0; i < 4; i++)
                {
                    if (Controls.Find("chkKouzou" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                    {
                        searchOptions.Add(id_list[c.Text]);
                    }
                }

                for (int i = 0; i < 5; i++)
                {
                    if (Controls.Find("chkIchi" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                    {
                        searchOptions.Add(id_list[c.Text]);
                    }
                }

                for (int i = 0; i < 4; i++)
                {
                    if (Controls.Find("chkBath" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                    {
                        searchOptions.Add(id_list[c.Text]);
                    }
                }

                for (int i = 0; i < 8; i++)
                {
                    if (Controls.Find("chkHoui" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                    {
                        searchOptions.Add(id_list[c.Text]);
                    }
                }

                for (int i = 0; i < 6; i++)
                {
                    if (Controls.Find("chkKitchen" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                    {
                        searchOptions.Add(id_list[c.Text]);
                    }
                }

                for (int i = 0; i < 8; i++)
                {
                    if (Controls.Find("chkNyukyo" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                    {
                        searchOptions.Add(id_list[c.Text]);
                    }
                }

                for (int i = 0; i < 5; i++)
                {
                    if (Controls.Find("chkSetsubi" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                    {
                        searchOptions.Add(id_list[c.Text]);
                    }
                }

                for (int i = 0; i < 8; i++)
                {
                    if (Controls.Find("chkShitunai" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                    {
                        searchOptions.Add(id_list[c.Text]);
                    }
                }

                for (int i = 0; i < 5; i++)
                {
                    if (Controls.Find("chkChusya" + (i + 1).ToString(), true)[0] is CheckBox c)
                        if (c != null && c.Checked)
                        {
                            searchOptions.Add(id_list[c.Text]);
                        }
                }

                for (int i = 0; i < 2; i++)
                {
                    if (Controls.Find("chkReidanbou" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                    {
                        searchOptions.Add(id_list[c.Text]);
                    }
                }

                for (int i = 0; i < 5; i++)
                {
                    if (Controls.Find("chkSecurity" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                    {
                        searchOptions.Add(id_list[c.Text]);
                    }
                }

                for (int i = 0; i < 2; i++)
                {
                    if (Controls.Find("chkJogai" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                    {
                        searchOptions.Add(id_list[c.Text]);
                    }
                }

                for (int i = 0; i < 7; i++)
                {
                    if (Controls.Find("chkSonota" + (i + 1).ToString(), true)[0] is CheckBox c && c.Checked)
                    {
                        searchOptions.Add(id_list[c.Text]);
                    }
                }

                // 対象件数を取得する
                var taisyo_kensu = SearchTaisyoKensu();

                if (taisyo_kensu <= 0)
                {
                    return;
                }

                // 対象を分割する

                var split_count = taisyo_kensu / 20 + 1;

                List<int> tmp = [];
                for (int i = 0; i < taisyo_kensu; i += 20)
                {
                    tmp.Add(i);
                }

                chunks = tmp.Select((v, i) => new { v, i })
                    .GroupBy(x => x.i / (1 + tmp.Count / THREAD_COUNT))
                    .Select(g => g.Select(x => x.v)).ToList();

                // パラメータを取得
                param = new()
                {
                    chomeiIds = Array.Empty<string>(),
                    cityIds = Array.Empty<string>(),
                    commutes =
                    [
                        new()
                        {
                            sourceStationId = STATION_SHIBUYA,
                            timeMinutes = int.Parse(cmbToShibuya1.Text.Replace("分", "")),
                            changeCount = new Changecount()
                            {
                                value = cmbToShibuya2.SelectedIndex == 0 ? -1 : int.Parse(cmbToShibuya2.Text.Replace("回", "")),
                                hasValue = cmbToShibuya2.SelectedIndex != 0
                            }
                        },
                        new()
                        {
                            sourceStationId = STATION_AKIHABARA,
                            timeMinutes = int.Parse(cmbToAkiba1.Text.Replace("分", "")),
                            changeCount = new Changecount()
                            {
                                value = cmbToAkiba2.SelectedIndex == 0 ? -1 : int.Parse(cmbToAkiba2.Text.Replace("回", "")),
                                hasValue = cmbToAkiba2.SelectedIndex != 0
                            }
                        }
                    ],
                    duringMax = new Duringmax()
                    {
                        hasValue = cmbToho.SelectedIndex != 0,
                        value = cmbToho.SelectedIndex == 0 ? 0 : int.Parse(cmbToho.Text.Replace("分以下", ""))
                    },
                    includeAdminFee = chkChinryo3.Checked,
                    isNewArrival = false,
                    keywords = Array.Empty<string>(),
                    layoutNames = [.. layoutNames],
                    oldMax = new Oldmax()
                    {
                        hasValue = cmbNensu.SelectedIndex != 0,
                        value = cmbNensu.SelectedIndex == 0 ? 0 : int.Parse(cmbNensu.Text.Replace("築", "").Replace("年以下", ""))
                    },
                    prefectureIds = [.. prefectureIds],
                    rentMax = new Rentmax()
                    {
                        hasValue = cmbChinryo2.SelectedIndex != 0,
                        value = cmbChinryo2.SelectedIndex == 0 ? 0 : int.Parse(cmbChinryo2.Text.Replace("万以下", "")) * 10000
                    },
                    rentMin = new Rentmin()
                    {
                        hasValue = cmbChinryo1.SelectedIndex != 0,
                        value = cmbChinryo1.SelectedIndex == 0 ? 0 : int.Parse(cmbChinryo1.Text.Replace("万以上", "")) * 10000
                    },
                    searchOptionIds = [.. searchOptions],
                    shatakuCode = new Shatakucode()
                    {
                        hasValue = false,
                        value = ""
                    },
                    squareMax = new Squaremax()
                    {
                        hasValue = cmbSenyu2.SelectedIndex != 0,
                        value = cmbSenyu2.SelectedIndex == 0 ? 0 : int.Parse(cmbSenyu2.Text.Replace("㎡以下", ""))
                    },
                    squareMin = new Squaremin()
                    {
                        hasValue = cmbSenyu1.SelectedIndex != 0,
                        value = cmbSenyu1.SelectedIndex == 0 ? 0 : int.Parse(cmbSenyu1.Text.Replace("㎡以上", ""))
                    },
                    stationIds = Array.Empty<string>(),
                    limit = 20,
                    listType = 2,
                    offset = new Offset()
                    {
                        hasValue = true,
                        value = "0"
                    },
                    searchPurpose = true,
                    searchSessionId = string.Empty,
                    sortType = 1
                };

            });

            if (chunks.Count <= 0)
            {
                return null;
            }

            // 対象の部屋が存在しなくなるまでループ
            List<Task<Tuple<List<JsonNode>, List<JsonNode>>>> task_list = [];

            for (int i = 0; i < chunks.Count; i++)
            {
                // オフセットリスト
                var offset_list = chunks[i];

                // パラメータをコピー
                Parameter2 param_tmp = new(param);

                var index = i;

                task_list.Add(Task.Run(() =>
                {
                    List<JsonNode> result1 = [];
                    List<JsonNode> result2 = [];

                    for (int j = 0; j < offset_list.Count(); j++)
                    {
                        Debug.WriteLine("スレッド番号 = " + (index + 1).ToString() + " : オフセット = " + offset_list.ToArray()[j].ToString());

                        // 個別のパラメータを指定
                        param_tmp.offset = new Offset()
                        {
                            hasValue = true,
                            value = offset_list.ToArray()[j].ToString()
                        };

                        using var client = new HttpClient();

                        // APIを利用して検索する
                        var task = client.PostAsync("https://api.user.canary-app.com/v1/chintaiRooms:search", new StringContent(JsonSerializer.Serialize(param_tmp), Encoding.UTF8, "application/json"));
                        task.Wait();

                        // HttpContentから文字列を取得する
                        var task2 = task.Result.Content.ReadAsStringAsync();
                        task2.Wait();

                        // Jsonオブジェクトにパースする
                        var json_data = JsonNode.Parse(task2.Result);

                        // 配列長を取得
                        var length1 = json_data["chintaiEstates"].AsArray().Count;
                        var length2 = json_data["chintaiRooms"].AsArray().Count;

                        if (length1 <= 0 || length2 <= 0)
                        {
                            Debug.WriteLine("スレッド番号 = " + (index + 1).ToString() + " : ループをBREAK");
                            break;
                        }

                        for (int k = 0; k < length1; k++)
                        {
                            result1.Add(json_data["chintaiEstates"][k]);
                        }

                        for (int k = 0; k < length2; k++)
                        {
                            result2.Add(json_data["chintaiRooms"][k]);
                        }
                    }

                    return new Tuple<List<JsonNode>, List<JsonNode>>(result1, result2);
                }));
            }

            // すべてのスレッドが完了するまで待機する
            for (int i = 0; i < task_list.Count; i++)
            {
                task_list[i].Wait();
                result_estates.AddRange(task_list[i].Result.Item1);
                result_rooms.AddRange(task_list[i].Result.Item2);
            }

            // 計測停止
            sw.Stop();

            TimeSpan ts = sw.Elapsed;
            Debug.WriteLine($"　{ts.Hours}時間 {ts.Minutes}分 {ts.Seconds}秒 {ts.Milliseconds}ミリ秒");
            Debug.WriteLine(result_estates.Count.ToString() + "件");

            // オリジナル条件を反映する
            List<JsonNode> extract_rooms = new();
            for (int i = 0; i < result_rooms.Count; i++)
            {
                if (chkOriginal1.Checked && result_rooms[i]["layoutDetail"]["value"].ToString().Contains("和"))
                {
                    continue;
                }

                extract_rooms.Add(result_rooms[i]);
            }

            return new Tuple<List<JsonNode>, List<JsonNode>>(result_estates, extract_rooms);
        }

        /// <summary>
        /// 賃貸情報をプレビュー表示する
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnPreview_Click(object sender, EventArgs e)
        {
            // 賃貸リストを生成する
            var result = GenerateChintaiList();

            if (result == null)
            {
                return;
            }

            // エクセルに落とし込む
            SaveChintaiData(result.Item1, result.Item2, true);
        }
    }
}
