# 開発環境での動作の仕方
# WindowsでPowerShellを起動（Virual BoxのUbuntu環境では動作しない）
# D:\vagrant\rpaに移動
# ruby rpa_seikyu.rb 引数    - 請求システムのRPA
# ruby rpa_rakuraku.rb 引数  - 楽楽販売のRPA

require 'selenium-webdriver'
require 'logger'
require 'yaml'
require 'csv'

# ログイン 処理
class SESSIONS
    class << self

        attr_reader :syori_cnt
        attr_writer :syori_cnt

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
                            msg  = "\n"
                            msg << "  a [-a]    Access連携出力（請求/施設）をCSVする処理を実行します \n"
                            msg << "  s [-s]    Access連携出力（請求）をCSVする処理を実行します \n"
                            msg << "  k [-k]    管理部提出データ1～3をCSV出力する処理を実行します \n"
                            msg << "  o [-o]    障害テーブル（RAG連携）をCSV出力する処理を実行します \n"
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
                        # 障害管理
                        when "o", "-o"
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
        def proc_init(argv)
            begin
                FileUtils.rm("./production.log", force: true)               # Logファイルの削除
                $logger = Logger.new('production.log')                      # Logの設定
                $logger.info("処理を開始しました。 引数: #{argv[0]}")
                $has_local = YAML.load_file("./local.yaml")                 # YAMLファイル読み込み
                return nil
            rescue => ex
                return "method - " + __method__.to_s + " : " + ex.message + ": YAMLファイルの読み込みに失敗しました。"
            end
        end

        # 終了処理
        def proc_end(driver, cat)
            
            if cat.nil?
                $logger.info("処理が正常終了しました。")
            else
                $logger.error("#{cat}")
                $logger.info("処理が異常終了しました。")
                driver.quit if !driver.nil?
            end
            $logger.close
        end

        # ログイン処理
        def proc_main(driver)
            
            @syori_cnt = 1
            
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
                
                # ログイン判断
                if driver.find_elements(class: 'fw-keyword').size >= 1
                    return "ログインID、または、パスワードに誤りがあります。"
                end

                $logger.info("#{@syori_cnt.to_s.rjust(2)} ログインができました。")
                return nil
            rescue => ex
                return "method - " + __method__.to_s + " : " + ex.message + ": 「ログイン処理」でエラーが発生しました。"
            end
        end
    end
end

# 楽楽販売の処理（請求処理）
class RAKURAKU
    class << self
        attr_reader :csv_arry
        attr_writer :csv_arry

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

                SESSIONS::syori_cnt += 1
                $logger.info("#{SESSIONS::syori_cnt.to_s.rjust(2)} 請求処理が選択されました。")
                return nil
            rescue => ex
                return "method - " + __method__.to_s + " : " + ex.message + ": 「請求処理」でエラーが発生しました。"
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
                        msg = "請求テーブルが選択されました。"
                    when "sisetu"
                        # 施設テーブルをクリック
                        driver.find_element(id: "nav-db-101163").click
                        sleep(0.3)
                        msg = "施設テーブルが選択されました。"
                    else
                        raise
                end
                
                SESSIONS::syori_cnt += 1
                $logger.info("#{SESSIONS::syori_cnt.to_s.rjust(2)} #{msg}")
                return nil
            rescue => ex
                return "method - " + __method__.to_s + " : " + ex.message + ": 「請求処理のテーブルの選択」でエラーが発生しました。"
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
                    
                    SESSIONS::syori_cnt += 1
                    $logger.info("#{SESSIONS::syori_cnt.to_s.rjust(2)} #{res[:menu]}が選択されました。")

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
                    
                    # データ件数が多くなってダウンロード時間が長くなれば、sleepは長くすること（特に施設）
                    minutes = (res[:tbl] = "sisetu" and res[:menu] = "Access連携出力") ? 12.0 : 5.5
                    
                    # ダウンロード
                    driver.find_element(id: "csv_confirm_start").click
                    sleep(minutes)
                    
                    # ダウンロードファイル
                    driver.find_element(id: "csv_complete_link").click
                    sleep(2.5)

                    # ダウンロードするCSVファイル名を取得
                    @csv_arry << driver.find_element(id: "csv_complete_link").text

                    # 閉じる
                    driver.find_element(id: "csv_complete_close").click
                    sleep(0.5)

                    SESSIONS::syori_cnt += 1
                    $logger.info("#{SESSIONS::syori_cnt.to_s.rjust(2)} 「#{@csv_arry[-1]}」がダウンロードできました。")

                    # メイン画面に戻る
                    driver.switch_to.default_content
                    sleep(0.3)

                    # iframeを取得
                    driver.switch_to.frame 'side'
                    sleep(0.3)
                end
                
                return nil
            rescue => ex
                return "method - " + __method__.to_s + " : " + ex.message + ": 「請求処理－各テーブル－各メニュー」でエラーが発生しました。"
            end
        end
    end
end

# 楽楽販売の処理（障害管理）
class SHOUGAIK
    class << self
        attr_reader :csv_arry
        attr_writer :csv_arry

        # 障害管理
        def proc_main(driver)
            begin
                # iframeを取得
                driver.switch_to.frame 'side'
                sleep(0.3)

                # 障害管理・導入後の対応を選択
                driver.find_element(id: "nav-dbg-100170").click
                sleep(0.3)

                # 幅を大きくする
                driver.manage.window.resize_to(1200, 800)
                sleep(0.3)

                SESSIONS::syori_cnt += 1
                $logger.info("#{SESSIONS::syori_cnt.to_s.rjust(2)} 障害管理・導入後の対応が選択されました。")

                return nil
            rescue => ex
                return "method - " + __method__.to_s + " : " + ex.message + ": 「障害管理」でエラーが発生しました。"
            end
        end

        # 障害管理－テーブルを選択
        def proc_table(driver, table)
            begin
                case table
                    when "shogai"
                        # 障害・個別対応テーブルを選択
                        driver.find_element(id: "nav-db-101344").click
                        sleep(0.3)
                        msg = "障害・個別対応テーブルが選択されました。"
                    else
                        raise
                end
                
                SESSIONS::syori_cnt += 1
                $logger.info("#{SESSIONS::syori_cnt.to_s.rjust(2)} #{msg}")
                return nil
            rescue => ex
                return "method - " + __method__.to_s + " : " + ex.message + ": 「障害管理のテーブルの選択」でエラーが発生しました。"
            end
        end

        # 障害管理－各テーブル－各メニュー
        def proc_syori(driver, tbl_id, syori_kbn)
            begin
                menu_arry = case syori_kbn
                    when "o"
                        [{tbl: "shogai", menu: "障害テーブル（RAG連携）"}]
                end

                @csv_arry = []

                # デフォルト画面に戻る
                driver.switch_to.default_content
                sleep(0.3)

                # メイン画面に遷移
                driver.switch_to.frame driver.find_element(name: 'main')
                sleep(0.3)

                menu_arry.each do |res|
                    
                    # 一覧画面
                    select = Selenium::WebDriver::Support::Select.new(driver.find_element(name: 'listFormatId'))
                    sleep(0.3)
                    
                    # 一覧画面－障害テーブル（RAG連携）を選択
                    select.select_by(:text, "#{res[:menu]}")
                    sleep(3.0)
                    
                    SESSIONS::syori_cnt += 1
                    $logger.info("#{SESSIONS::syori_cnt.to_s.rjust(2)} #{res[:menu]}が選択されました。")

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
                    driver.find_element(id: "csv_confirm_start").click
                    sleep(2.5)

                    # ダウンロードファイル
                    driver.find_element(id: "csv_complete_link").click
                    sleep(2.0)

                    # ダウンロードするCSVファイル名を取得
                    @csv_arry << driver.find_element(id: "csv_complete_link").text

                    # 閉じる
                    driver.find_element(id: "csv_complete_close").click
                    sleep(0.5)

                    SESSIONS::syori_cnt += 1
                    $logger.info("#{SESSIONS::syori_cnt.to_s.rjust(2)} 「#{@csv_arry[-1]}」がダウンロードできました。")

                    # メイン画面に戻る
                    driver.switch_to.default_content
                    sleep(0.3)
                end

                return nil
            rescue => ex
                return "method - " + __method__.to_s + " : " + ex.message + ": 「請求処理－各テーブル－各メニュー」でエラーが発生しました。"
            end
        end
    end
end

# CSVファイルの操作
class CSVFILES
    class << self

        # CSVファイルのリネーム
        def change_csv(csv_arry)

            msg = nil
            csv_arry.each do |csv|

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
                sleep(0.5)

                SESSIONS::syori_cnt += 1
                $logger.info("#{SESSIONS::syori_cnt.to_s.rjust(2)} 「#{csv}」→「#{csv[0..-21]}.csv」にリネームしました。")
            end
            return msg
        end

        # CSVファイルをJSONファイルに変換
        def create_json(file)
            begin
                return "「#{file}.json」が存在しません。" if !File.exist?("#{file}.json")
                
                header = ["ID","ユーザー","内容","現象／原因","処置","備考","連絡受付日","ユーザーキー"]
                cnt = 0
                
                fil = File.open("#{file}.json", "a")

                CSV.foreach("#{file}.csv") do |csv|
                
                    cnt += 1
                    next if cnt == 1
                    
                    hash_line, user = {}, ""
            
                    header.length.times do |idx|
                        case idx
                            when 0
                                hash_line["#{header[idx]}"] = csv[idx] 
                            when 1
                                user = csv[idx].to_s
                            when 2
                                # ユーザー名とエンドユーザーをユーザーの１つにまとめる
                                hash_line["#{header[idx-1]}"] = user == "" ? csv[idx] : user
                            when 2..(header.length)
                                hash_line["#{header[idx-1]}"] = csv[idx] 
                        end
                    end

                    fil.write("#{hash_line}\n")
                    # break if cnt == 3
                end

                fil.close

                SESSIONS::syori_cnt += 1
                $logger.info("#{SESSIONS::syori_cnt.to_s.rjust(2)} 「#{File.basename(file)}.csv」→「#{File.basename(file)}.json」への書き込みができました。")
                return nil

            rescue => ex
                return "method - " + __method__.to_s + " : " + ex.message + ": JSONファイルの作成に失敗しました。"
            end
        end
    end
end

# 共通クラス
class COMMONCL
    class << self
        # 空ファイルを作成
        def create_file(file)
            begin
                File.delete(file) if File.exist?(file)
                fil = File.open(file, "w")
                fil.close
                return nil
            rescue => ex
                return "method - " + __method__.to_s + " : " + ex.message + ": 「#{file}」の空ファイルの作成でエラーが発生しました。"
            end
        end

        # JSONファイルをリネーム
        def create_bkup(file)
            begin
                return "#{file}が存在しません。" if !File.exist?(file)
                File.delete("#{file}.bkup")    if File.exist?("#{file}.bkup")
                File.rename(file, "#{file}.bkup")

                SESSIONS::syori_cnt += 1
                $logger.info("#{SESSIONS::syori_cnt.to_s.rjust(2)} 「#{File.basename(file)}」に「.bkup」を付加してリネームしました。")

                return nil
            rescue => ex
                return "method - " + __method__.to_s + " : " + ex.message + ": 「#{file}」の.bkupへのリネームでエラーが発生しました。"
            end
        end
        
        # 改行コードを。に変換
        def change_kaigyo(file)
            begin
                filout = File.open("#{file}.bkup", "r")
                filin  = File.open(file, "a")

                filout.each_with_index do |row, idx|
                    row.gsub!(/(\\r\\n)+/, "。") if row.include?("\\r\\n")
                    filin.write(row)
                    # break if idx == 1
                end

                filout.close
                SESSIONS::syori_cnt += 1
                $logger.info("#{SESSIONS::syori_cnt.to_s.rjust(2)} 「#{File.basename(file)}」で「\\r\\n → 。」に変換しました。")
                return nil
            rescue => ex
                return "method - " + __method__.to_s + " : " + ex.message + ": 「#{file}」の改行コードの変換でエラーが発生しました。"
            end
        end
    end
end

# 開発テスト用
class TEST_DEV
    class << self
        # ソースを標準出力する
        def disp_source(driver)
            puts driver.page_source
        end
    end
end

# ------------------------------------------------------------------------------
# メイン処理
# ------------------------------------------------------------------------------

ret, syori_kbn, driver = nil, nil, nil

cat = catch(:goto_err) do
    
    # 初期処理
    ret = SESSIONS.proc_init(ARGV)
    throw :goto_err, ret if !ret.nil?

    # 引数の入力チェック
    ret = SESSIONS.check_init(ARGV)
    throw :goto_err, ret if !(ret.length == 1)
    syori_kbn = ret

    # 起動ブラウザの設定
    options = Selenium::WebDriver::Chrome::Options.new
    options.detach = true
    options.add_argument('--log-level=1')
    options.exclude_switches << 'enable-logging'
    
    driver = Selenium::WebDriver.for :chrome, options: options

    # ログイン処理
    ret = SESSIONS.proc_main(driver)
    throw :goto_err, ret if !ret.nil?
        
    case syori_kbn
        when "l"
        when "a", "s", "k"
            
            # 請求処理
            ret = RAKURAKU.proc_main(driver)
            throw :goto_err, ret if !ret.nil?
            
            # 処理ルーティン
            syori_arry = [{tbl_nm: "請求", tbl_id: "seikyu"}, 
                          {tbl_nm: "施設", tbl_id: "sisetu"}]

            syori_arry.each do |res|

                # 請求処理－各テーブル
                ret = RAKURAKU.proc_table(driver, res[:tbl_id])
                throw :goto_err, ret if !ret.nil?

                # 請求処理－各テーブル－各メニュー
                ret = RAKURAKU.proc_syori(driver, res[:tbl_id], syori_kbn)
                throw :goto_err, ret if !ret.nil?

                # CSVファイルのリネーム
                ret = CSVFILES.change_csv(RAKURAKU::csv_arry)
                throw :goto_err, ret if !ret.nil?
            end

        when "o"
            # 障害管理
            ret = SHOUGAIK.proc_main(driver)
            throw :goto_err, ret if !ret.nil?

            # 処理ルーティン
            syori_arry = [{tbl_nm: "障害", tbl_id: "shogai"}]
                         
            syori_arry.each do |res|
                
                # 障害管理－各テーブル
                ret = SHOUGAIK.proc_table(driver, res[:tbl_id])
                throw :goto_err, ret if !ret.nil?

                # 障害管理－各テーブル－各メニュー
                ret = SHOUGAIK.proc_syori(driver, res[:tbl_id], syori_kbn)
                throw :goto_err, ret if !ret.nil?

                # CSVファイルのリネーム
                ret = CSVFILES.change_csv(SHOUGAIK::csv_arry)
                throw :goto_err, ret if !ret.nil?

                file = "#{$has_local["csv_path"]}/障害・個別対応テーブル：障害テーブル一覧"
                
                # 空ファイルを作成
                ret = COMMONCL.create_file("#{file}.json")
                throw :goto_err, ret if !ret.nil?

                # CSVファイルをJSONファイルに変換
                ret = CSVFILES.create_json(file)
                throw :goto_err, ret if !ret.nil?

                # バックアップファイルにリネーム
                ret = COMMONCL.create_bkup("#{file}.json")
                throw :goto_err, ret if !ret.nil?

                # 空ファイルを作成
                ret = COMMONCL.create_file("#{file}.json")
                throw :goto_err, ret if !ret.nil?
                
                # 改行コードを変換
                ret = COMMONCL.change_kaigyo("#{file}.json")
                throw :goto_err, ret if !ret.nil?
            end
    end
    throw :goto_err, nil
end

# 終了処理
SESSIONS.proc_end(driver, cat)
