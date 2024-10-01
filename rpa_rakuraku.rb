# 開発環境での動作の仕方
# WindowsでPowerShellを起動
# D:\vagrant\rpaに移動
# ruby rpa_rakuraku.rb k 管理部提出データ出力１～3をCSV出力する場合
# ruby rpa_rakuraku.rb s Access連携出力（請求）をCSV出力する場合
# ruby rpa_rakuraku.rb a Access連携出力（請求/施設）をCSV出力する場合
# ruby rpa_rakuraku.rb l 楽楽販売にログインするだけの場合

require 'selenium-webdriver'
require 'logger'
require 'yaml'

# ログイン 処理
class SESSIONS
    class << self

        # 引数チェック
        def check_init(argv)
            
            msg = nil
            begin
                if argv.size <= 0
                    msg = "引数が設定されていません。"
                elsif argv.size >= 2
                    msg = "引数は１つしか設定できません。"
                else
                    case argv[0].to_s
                        # ヘルプ
                        when "h", "-h"
                            msg  = "  a [-a]    Access連携出力（請求/施設）をCSVする処理を実行します \n"
                            msg  = "  s [-s]    Access連携出力（請求）をCSVする処理を実行します \n"
                            msg << "  k [-k]    管理部提出データ1～3をCSV出力する処理を実行します \n"
                            msg << "  l [-l]    ログインだけをおこないます \n"
                            msg << "  h [-h]    ヘルプを表示します"
                        # Access連携出力（請求/施設）
                        when "a", "-a"
                            msg = argv[0][-1].to_s
                        # Access連携出力（請求）
                        when "s", "-s"
                            msg = argv[0][-1].to_s
                        # 管理部提出データ1～3
                        when "k", "-k"
                            msg = argv[0][-1].to_s
                        # ログイン
                        when "l", "-l"
                            msg = argv[0][-1].to_s
                        else
                            msg = "引数が正しく指定されていません。"
                    end
                end

                return msg
            rescue => ex
                msg = "method - " + __method__.to_s + " : " + ex.message + ": 引数チェックでエラーが発生しました。"
                $logger.error("#{msg}")
                return msg
            end
        end

        # 初期処理
        def proc_init
            begin
                FileUtils.rm("./production.log", force: true)               # Logファイルの削除
                $has_local = YAML.load_file("./local.yaml")                 # YAMLファイル読み込み
                $logger = Logger.new('production.log')                      # Logの設定
                $logger.info("処理を開始しました。")

                return nil
            rescue => ex
                msg = "method - " + __method__.to_s + " : " + ex.message + ": YAMLファイルの読み込みに失敗しました。"
                $logger.error("#{msg}")
                return msg
            end
        end

        # 終了処理
        def proc_end(driver, ret, cat)
            
            if cat.nil?
                $logger.info("処理が正常終了しました。")
            elsif cat.is_a?(Array)
                case cat[1]
                    when nil
                        $logger.error("#{ret}")
                        $logger.error("処理が異常終了しました。")
                    when "chk"
                        $logger.error("#{ret}")
                    else
                        $logger.error("不明なエラーです。")
                end
                driver.quit
            else
                $logger.error("不明なエラーです。")
            end
            $logger.close
        end

        # ログイン処理
        def proc_main(driver)

            begin
                wait = Selenium::WebDriver::Wait.new(timeout: 2)
                driver.navigate.to 'https://hncapitol.rakurakuhanbai.jp/wfecn6a/'
                sleep(0.5)

                # ログインID
                driver.find_element(name: 'loginId').send_keys "#{$has_local["rakrak"]["id"]}"
                sleep(0.2)

                # パスワード
                driver.find_element(name: 'loginPassword').send_keys "#{$has_local["rakrak"]["pw"]}"
                sleep(0.2)

                # ログインボタン
                driver.find_element(id: 'jq-loginSubmit').click
                sleep(0.3)

                $logger.info("ログインができました。")
                return nil
            rescue => ex
                msg = "method - " + __method__.to_s + " : " + ex.message + ": 「ログイン処理」でエラーが発生しました。"
                $logger.error("#{msg}")
                return msg
            end
        end
    end
end

# 楽楽販売の処理
class RAKURAKU
    class << self
        attr_reader :csv_arry

        # 請求処理
        def proc_main(driver)
            begin
                # iframeを取得
                driver.switch_to.frame 'side'
                sleep(0.3)

                # 請求処理をクリック
                driver.find_element(id: "nav-dbg-100136").click
                sleep(0.3)

                # 幅を大きくする
                driver.manage.window.resize_to(1200, 800)
                sleep(0.3)

                $logger.info("請求処理が選択されました。")
                return nil
            rescue => ex
                msg = "method - " + __method__.to_s + " : " + ex.message + ": 「請求処理」でエラーが発生しました。"
                $logger.error("#{msg}")
                return msg
            end
        end

        # 請求処理－テーブルを選択
        def proc_table(driver, table)
            begin
                case table
                    when "seikyu"
                        # 請求テーブルをクリック
                        driver.find_element(id: "nav-db-101164").click
                        sleep(0.3)
                        $logger.info("請求テーブルが選択されました。")
                    when "sisetu"
                        # 施設テーブルをクリック
                        driver.find_element(id: "nav-db-101163").click
                        sleep(0.3)
                        $logger.info("施設テーブルが選択されました。")
                    else
                        raise
                end
                
                return nil
            rescue => ex
                msg = "method - " + __method__.to_s + " : " + ex.message + ": 「請求処理のテーブルの選択」でエラーが発生しました。"
                $logger.error("#{msg}")
                return msg
            end
        end


        # 請求処理－各テーブル－各メニュー
        def proc_syori(driver, tbl_id, syori_kbn)
            begin
                menu_arry = case syori_kbn
                    when "a"
                        [{tbl: "seikyu", menu: "Access連携出力", xpath: "//*[@id='menuli_102226']/a/div"},
                         {tbl: "sisetu", menu: "Access連携出力", xpath: "//*[@id='menuli_102227']/a/div"}]
                    when "s"
                        [{tbl: "seikyu", menu: "Access連携出力", xpath: "//*[@id='menuli_102226']/a/div"}]
                    when "k"
                        [{tbl: "sisetu", menu: "管理部提出データ出力１", xpath: "//*[@id='menuli_102231']/a/div"}, 
                         {tbl: "sisetu", menu: "管理部提出データ出力２", xpath: "//*[@id='menuli_102232']/a/div"}, 
                         {tbl: "sisetu", menu: "管理部提出データ出力３", xpath: "//*[@id='menuli_102603']/a/div"}]
                    else
                        raise
                end
                
                @csv_arry = []

                menu_arry.each do |res|
                    
                    next if res[:tbl] != tbl_id

                    # 各処理メニューを選択
                    driver.find_element(xpath: "#{res[:xpath]}").click
                    sleep(0.3)
                    
                    # メイン画面に戻る
                    driver.switch_to.default_content
                    sleep(0.3)
                    
                    # iframeを取得
                    driver.switch_to.frame 'main'
                    sleep(0.3)

                    # ハンバーガーメニュー
                    driver.find_element(id: "link_menu_box").click
                    sleep(0.5)

                    # CSV出力
                    driver.find_element(id: "popupCsvExport").click
                    sleep(0.5)

                    # UTF-8で出力する
                    driver.find_element(id: "csv_downloadUtf8").click
                    sleep(0.5)

                    # ダウンロード
                    # データ件数が多くなればダウンロード時間が長くなるため、sleepは都度、メンテする必要がある
                    driver.find_element(id: "csv_confirm_start").click
                    sleep(4)

                    # ダウンロードファイル
                    driver.find_element(id: "csv_complete_link").click
                    sleep(3)

                    # ダウンロードするCSVファイル名を取得
                    @csv_arry << driver.find_element(id: "csv_complete_link").text

                    # 閉じる
                    driver.find_element(id: "csv_complete_close").click
                    sleep(0.5)

                    $logger.info("「#{@csv_arry[-1]}」がダウンロードできました。")

                    # メイン画面に戻る
                    driver.switch_to.default_content
                    sleep(0.3)

                    # iframeを取得
                    driver.switch_to.frame 'side'
                    sleep(0.3)
                end
                
                return nil
            rescue => ex
                msg = "method - " + __method__.to_s + " : " + ex.message + ": 「請求処理」でエラーが発生しました。"
                $logger.error("#{msg}")
                return msg
            end
        end
    end
end

# CSVファイルの操作
class CSVFILES
    class << self

        # CSVファイルのリネーム
        def change_csv

            msg = nil
            RAKURAKU::csv_arry.each do |csv|

                moto = "#{$has_local["csv_path"]}/#{csv}"                       # CSVファイルのリネーム元
                saki = "#{$has_local["csv_path"]}/#{csv[0..-21]}" + ".csv"      # CSVファイルのリネーム先
                
                # ファイルの存在チェック
                if !File.exist?(moto)
                    msg = "CSVファイルのダウンロード先のパスが存在しません。"
                    break
                end

                # 既存ファイルの削除
                FileUtils.rm(saki, force: true)
                
                # ファイルのリネーム
                FileUtils.mv(moto, saki)

                $logger.info("「#{csv}」→「#{csv[0..-21]}.csv」にリネームしました。")
            end
            return msg
        end
    end
end

# ------------------------------------------------------------------------------
# メイン処理
# ------------------------------------------------------------------------------

ret, syori_kbn, driver = nil, nil, nil

cat = catch(:goto_err) do

    # 引数の入力チェック
    ret = SESSIONS.check_init(ARGV)
    throw :goto_err, [ret, "chk"] if !(ret.length == 1)
    syori_kbn = ret

    # 起動ブラウザの設定
    options = Selenium::WebDriver::Chrome::Options.new
    options.detach = true
    driver = Selenium::WebDriver.for :chrome, options: options

    # 初期処理
    ret = SESSIONS.proc_init
    throw :goto_err, [ret, nil] if !ret.nil?

    # ログイン処理
    ret = SESSIONS.proc_main(driver)
    throw :goto_err, [ret, nil] if !ret.nil?
    
    # 請求処理
    if ( syori_kbn == "a" or syori_kbn == "s" or syori_kbn == "k" )
        ret = RAKURAKU.proc_main(driver)
        throw :goto_err, [ret, nil] if !ret.nil?
    end

    # 処理ルーティン
    syori_arry = [{tbl_nm: "請求", tbl_id: "seikyu", syori_kbn: ["a", "s"]}, 
                  {tbl_nm: "施設", tbl_id: "sisetu", syori_kbn: ["a", "k"]}
                 ]

    syori_arry.each do |res|

        if res[:syori_kbn].include?(syori_kbn)
            
            # 請求処理－各テーブル
            ret = RAKURAKU.proc_table(driver, res[:tbl_id])
            throw :goto_err, [ret, nil] if !ret.nil?
            
            # 請求処理－各テーブル－各メニュー
            ret = RAKURAKU.proc_syori(driver, res[:tbl_id], syori_kbn)
            throw :goto_err, [ret, nil] if !ret.nil?

            # CSVファイルのリネーム
            ret = CSVFILES.change_csv
            throw :goto_err, [ret, nil] if !ret.nil?
        end
    end
    throw :goto_err, nil
end

# 終了処理
SESSIONS.proc_end(driver, ret, cat)
