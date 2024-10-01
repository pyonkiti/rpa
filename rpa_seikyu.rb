# 開発環境での動作の仕方
# WindowsでPowerShellを起動
# D:\vagrant\rpaに移動
# ruby rpa_seikyu.rb u Sofinet Cloudの連携（ユーザー）
# ruby rpa_seikyu.rb s Sofinet Cloudの連携（ユーザー/施設）
# ruby rpa_seikyu.rb n 楽楽のCSVインポートのインポート（入金仕入出力）
# ruby rpa_seikyu.rb c 請求月計算

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
                            msg  = "  u [-u]    Sofinet Cloudの連携（ユーザー） \n"
                            msg  = "  s [-s]    Sofinet Cloudの連携（ユーザー/施設） \n"
                            msg << "  n [-n]    楽楽のCSVインポートのインポート（入金仕入出力） \n"
                            msg << "  c [-c]    請求月計算 \n"
                            msg << "  h [-h]    ヘルプを表示します"
                        # Sofinet Cloudの連携（ユーザー）
                        when "u", "-u"
                            msg = argv[0][-1].to_s
                        # Sofinet Cloudの連携（施設）
                        when "s", "-s"
                            msg = argv[0][-1].to_s
                        # 管理部提出データ1～3
                        when "n", "-n"
                            msg = argv[0][-1].to_s
                        # 請求月計算
                        when "c", "-c"
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

        # 初期 処理
        def proc_init
            begin
                FileUtils.rm("./production.log", force: true)               # Logファイルの削除
                $has_local = YAML.load_file("./local.yaml")                 # YAMLファイル読み込み
                $logger = Logger.new('production.log')                      # Logの設定
                $logger.info("処理を開始しました。")

                return nil
            rescue => ex
                msg = "method - " + __method__.to_s + " : " + ex.message + ": YAMLファイルの読み込みに失敗しました。"
                $logger.warn("#{msg}")
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

        # ログイン 処理
        def proc_main(driver)
            begin
                svr_cont = case $has_local["svr_cont"]
                    when "dev" then "http://192.168.33.10:3000/"
                    when "pro" then "http://192.168.19.11:8000/"
                    else 
                        return "接続先のURLが正しく設定されていません。"
                end
                
                driver.navigate.to "#{svr_cont}"
                
                # ログインID
                driver.find_element(id: 'session_name_id').send_keys "#{$has_local["seikyu"]["id"]}"
                sleep(0.3)

                # パスワード
                driver.find_element(id: 'session_password').send_keys "#{$has_local["seikyu"]["pw"]}"
                sleep(0.3)

                # ログインボタン
                driver.find_element(name: 'button').click
                sleep(0.5)

                # ログインチェック
                if driver.find_element(class: 'alert').text != "ログインしました。"
                    return "ログインID、もしくは、パスワードに誤りがあります。"
                end

                $logger.info("ログインができました。")
                return nil
            rescue => ex
                msg = "method - " + __method__.to_s + " : " + ex.message + ": 「ログイン処理」でエラーが発生しました。"
                $logger.warn("#{msg}")
                return msg
            end
        end
    end
end

# -----------------------------------------------------------------------------

# Sofinet Cloudの連携
class RAKUREN_SCLOUD
    class << self

        # Sofinet Cloudの連携
        def proc_main(driver, syori_kbn)

            begin
                # 幅を大きくする
                driver.manage.window.resize_to(1200, 1020)
                sleep(0.3)

                # クラウド連携メニューを選択
                driver.find_element(xpath: "/html/body/div[1]/ul[1]/li[1]/a").click
                sleep(1.0)

                2.times do |cnt|

                    # CSVファイルのパスを指定
                    csv_file = case syori_kbn
                        # ユーザーの作成
                        when "u"
                            break if cnt == 1
                            "#{$has_local["csv_path"]}/請求テーブル：Access連携出力.csv"
                        # 施設の作成
                        when "s"
                            if cnt == 0
                                "#{$has_local["csv_path"]}/請求テーブル：Access連携出力.csv"
                            else
                                "#{$has_local["csv_path"]}/施設テーブル：Access連携出力.csv"
                            end
                        else raise
                    end
                    
                    # ファイルの存在チェック
                    if !File.exist?(csv_file)
                        msg = "取り込む元のCSVファイル「#{csv_file}」が存在しません。"
                        return msg
                    end

                    # ファイル選択のID
                    id_tag_file = case syori_kbn
                        when "u" then "tag_file_seikyu"                                         # ユーザーの作成
                        when "s" then (cnt == 0) ? "tag_file_seikyu" : "tag_file_shisetu"       # 施設の作成
                        else raise
                    end

                    # ファイルを選択
                    driver.find_element(id: "#{id_tag_file}").send_keys csv_file
                    sleep(1.0)
                    
                    # ファイル選択のID
                    id_tag_submit, minutes, msg = case syori_kbn
                        when "u" then ["tag_submit_seikyu", 1.0, "処理１"]                                                         # ユーザーの作成
                        when "s" then (cnt == 0) ? ["tag_submit_seikyu", 1.0, "処理１"] : ["tag_submit_shisetu", 15.0, "処理２"]    # 施設の作成
                        else raise
                    end
                    
                    # 《メモ》 処理１ 00秒   処理２ 14秒
                    # 処理１～２
                    driver.find_element(id: "#{id_tag_submit}").click
                    sleep(minutes)
                    
                    $logger.info("#{msg}を実行しました。")

                    # ラジオボタンのID
                    id_chk = case syori_kbn
                        when "u" then "chk_syori_kbn1"                                          # ユーザーの作成
                        when "s" then (cnt == 0) ? "chk_syori_kbn1" : "chk_syori_kbn2"          # 施設の作成
                        else raise
                    end

                    # ラジオボタンを選択
                    driver.find_element(id: "#{id_chk}").click
                    sleep(0.5)

                    # 実行のsleep設定
                    minutes = case syori_kbn
                        when "u" then  100.0                                                    # ユーザーの作成
                        when "s" then (cnt == 0) ? 100.0 : 1.0                                  # 施設の作成
                        else raise
                    end
                    
                    # 《メモ》実行 処理１の時 1分28秒 処理２の時 0分00秒
                    # 実行
                    begin
                        driver.find_element(id: "btn_renkei_jikko").click
                        sleep(minutes)
                        $logger.info("実行をクリックしました。")
                    rescue => ex
                        if ex.class.to_s == "Net::ReadTimeout"
                            $logger.warn("Sofinet Cloudの連携の実行でタイムアウトが発生しました。")
                        end
                    end
                    
                    # Excelユーザ/Excel施設のID
                    id_lnk_cloud, msg = case syori_kbn
                        when "u" then ["lnk_cloud_renkeis1", "Excelユーザ"]                                                     # ユーザーの作成
                        when "s" then (cnt == 0) ? ["lnk_cloud_renkeis1", "Excelユーザ"] : ["lnk_cloud_renkeis2", "Excel施設"]  # 施設の作成
                        else raise
                    end

                    # Excelユーザ/Excel施設
                    driver.find_element(id: "#{id_lnk_cloud}").click
                    sleep(3.0)
                    $logger.info("#{msg}をクリックしました。")
                end

                return nil
            rescue => ex
                msg = "method - " + __method__.to_s + " : " + ex.message + ": 「Sofinet Cloudの連携」でエラーが発生しました。"
                $logger.warn("#{msg}")
                return msg
            end
        end
    end
end

# -----------------------------------------------------------------------------

# 楽楽のCSVインポート/入金仕入出力 処理
class RAKUREN_NYUSHI
    class << self

        # 楽楽のCSVインポート
        def proc_main(driver)
        
            begin
                # 幅を大きくする
                driver.manage.window.resize_to(1200, 1020)

                # 入金仕入メニューを選択
                driver.find_element(xpath: "/html/body/div[1]/ul[1]/li[2]/a").click
                sleep(1.0)
                
                3.times do |cnt|

                    # 《メモ》 インポート1～3、実行の経過時間（2024/9時点）、26秒 22秒 18秒 15秒
                    # データ件数が多くなって処理時間が長くなった場合は、sleep時間をメンテする必要がある

                    # CSVファイルのパスを指定
                    csv_file = "#{$has_local["csv_path"]}/施設テーブル：管理部提出データ出力#{cnt+1}.csv"

                    # ファイルの存在チェック
                    if !File.exist?(csv_file)
                        msg = "取り込む元のCSVファイル「#{csv_file}」が存在しません。"
                        return msg
                    end

                    # ファイルを選択
                    driver.find_element(id: "tag_nyushi_file#{cnt+1}").send_keys csv_file
                    sleep(2.0)

                    # インポート１～３
                    driver.find_element(id: "tag_nyushi_submit#{cnt+1}").click
                    sleep(30)

                    $logger.info("インポート#{cnt+1}をクリックしました。")
                end

                # 実行
                driver.find_element(id: "lnk_nyushi_jikko").click
                sleep(20)
                
                $logger.info("実行をクリックしました。")

                return nil
            rescue => ex
                msg = "method - " + __method__.to_s + " : " + ex.message + ": 「楽楽のCSVインポート」でエラーが発生しました。"
                $logger.warn("#{msg}")
                return msg
            end
        end

        # 入金一覧Excel ／ 仕入一覧CSV出力
        def proc_down(driver)
            begin

                # 《メモ》
                # Excelファイルの場合「安全でないダウンロードがブロックされました」のメッセージが表示される。
                # Seleniumを使っている場合、回避できないため、手動で「保存」をクリックして正式ダウンロードを行ってもらう
                # CSVファイルは正常にダウンロードできる。

                # 入金一覧Excel
                driver.find_element(id: "lnk_export_nyu").click
                sleep(4.0)
                
                $logger.info("入金一覧Excelをクリックしました。")

                # 仕入一覧CSV出力
                driver.find_element(id: "lnk_export_shi").click
                sleep(3.0)

                $logger.info("仕入一覧CSV出力をクリックしました。")

                # 画面を再読み込み
                driver.navigate.refresh

                return nil
            rescue => ex
                msg = "method - " + __method__.to_s + " : " + ex.message + ": 「入金一覧Excel／仕入一覧CSV出力」でエラーが発生しました。"
                $logger.warn("#{msg}")
                return msg
            end
        end
    end
end

# -----------------------------------------------------------------------------

# 請求月計算 ／ 請求予定額計算
class RAKUREN_SEIKYU
    class << self
        #  請求月計算
        def proc_main(driver)
            begin
                # 幅を大きくする
                driver.manage.window.resize_to(1200, 1020)

                # 請求月計算メニューを選択
                driver.find_element(xpath: "/html/body/div[1]/ul[1]/li[3]/a").click
                sleep(1.0)

                # 計算区分
                driver.find_element(id: "chk_seikyus_kbn1").click
                sleep(0.5)

                # 実行
                # 《メモ》 実行時間 6秒
                driver.find_element(id: "btn_seikyus_jikko").click
                sleep(10.0)
                
                $logger.info("実行をクリックしました。")

                # CSV出力
                driver.find_element(xpath: "//*[@id='lnk_seikyus_csv']").click
                sleep(3.0)

                $logger.info("CSV出力をクリックしました。")

                return nil
            rescue => ex
                msg = "method - " + __method__.to_s + " : " + ex.message + ": 「請求月計算」でエラーが発生しました。"
                $logger.warn("#{msg}")
                return msg
            end
        end
    end
end

# ------------------------------------------------------------------------------
# メイン処理
# ------------------------------------------------------------------------------

driver, ret = nil, nil

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

    # Sofinet Cloudの連携
    if ( syori_kbn == "u" or syori_kbn == "s" )
        ret = RAKUREN_SCLOUD.proc_main(driver, syori_kbn)
        throw :goto_err, [ret, nil] if !ret.nil?
    end

    # 楽楽のCSVインポート
    if ( syori_kbn == "n" or syori_kbn == "c" )
        ret = RAKUREN_NYUSHI.proc_main(driver)
        throw :goto_err, [ret, nil] if !ret.nil?
    end

    # 入金一覧Excel ／ 仕入一覧CSV出力
    if syori_kbn == "n"
        ret = RAKUREN_NYUSHI.proc_down(driver)
        throw :goto_err, [ret, nil] if !ret.nil?
    end

    # 請求月計算
    if syori_kbn == "c"
         ret = RAKUREN_SEIKYU.proc_main(driver)
         throw :goto_err, [ret, nil] if !ret.nil?
    end
    throw :goto_err, nil
end

# 終了処理
SESSIONS.proc_end(driver, ret, cat)
