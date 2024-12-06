# 斡旋手数料の請求書をExcelからPDFに変換します
# 《開発環境での動作の仕方》
#   WindowsでPowerShellを起動（Virual BoxのUbuntu環境では動作しない）
#   D:\vagrant\rpaに移動
#   ruby rpa_convpdf.rb 

require 'win32ole'
require 'hexapdf'
require 'logger'
require 'yaml'
require 'fileutils'
require 'date'

# 共通処理
class COMMON
    class << self
        
        attr_reader   :file_exl                                         # Excelファイル名
        attr_accessor :syori_cnt                                        # ログに出力するinfo件数

        # 初期値のセット
        def initialize
            @file_exl = "支払手数料明細表_#{Date.today.year.to_s}#{Date.today.month.to_s}.xlsx"
            @syori_cnt = 0
        end
        
        # 初期処理
        def proc_init
            begin
                FileUtils.rm("./production.log", force: true)
                $logger = Logger.new('production.log')
                $logger.info("処理を開始しました。")
                $has_local = YAML.load_file("./local.yaml")
                return nil
            rescue => ex
                return "method - " + __method__.to_s + " : " + ex.message + ": YAMLファイルの読み込みに失敗しました。"
            end
        end

        # 終了処理
        def proc_end(cat)
            if cat.nil?
                $logger.info("処理が正常終了しました。")
            else
                $logger.error("#{cat}")
                $logger.error("処理が異常終了しました。")
            end
            $logger.close
        end

        # ファイルの存在チェック
        def check_file
            begin
                # Excelファイルのフォルダ存在確認
                if !Dir.exist?($has_local["convpdf"]["path_exl"])
                    return "「#{$has_local["convpdf"]["path_exl"]}」フォルダが存在しません。"
                end

                # PDFファイルのフォルダ存在確認
                if !Dir.exist?($has_local["convpdf"]["path_pdf"])
                    return "「#{$has_local["convpdf"]["path_pdf"]}」フォルダが存在しません。"
                end
                
                # Excelファイルの存在確認
                if !File.exist?("#{$has_local["convpdf"]["path_exl"]}\\#{@file_exl}")
                    return "「#{$has_local["convpdf"]["path_exl"]}\\#{@file_exl}」ファイルが存在しません。"
                end

                hash_file = {mae: nil, ato: nil}
                
                # フォルダにあるファイルをすべて取得
                hash_file[:mae] = Dir::entries("#{$has_local["convpdf"]["path_pdf"]}")

                # 削除対象の要素を省く
                hash_file[:ato] = hash_file[:mae].reject { |item| !item.end_with?(".pdf") }

                # PDFファイルの削除
                if hash_file[:ato].size > 0
                    hash_file[:ato].each do |item|
                        FileUtils.rm("#{$has_local["convpdf"]["path_pdf"]}\\#{item}", force: true)
                    end
                    @syori_cnt += 1
                    $logger.info("#{@syori_cnt.to_s.rjust(2)} 「#{$has_local["convpdf"]["path_pdf"]}」から「#{hash_file[:ato].join(' ,')}」の#{hash_file[:ato].size}ファイルを削除しました。")
                end

                return nil
            rescue => ex
                return "method - " + __method__.to_s + " : " + ex.message + ": Excelファイルの存在チェックでエラーが発生しました。"
            end
        end
    end
end

# Excel処理
class EXCEL
    class << self

        # 初期値のセット
        def initialize
        end

        # Excelの起動
        def proc_init
            begin
                excel = WIN32OLE.new('Excel.Application')
                excel.visible = true
                excel.displayAlerts = false

                WIN32OLE.const_load(excel, Excel)
                workbook = excel.Workbooks.Open("#{$has_local["convpdf"]["path_exl"]}\\#{COMMON::file_exl}")
                
                COMMON::syori_cnt += 1
                $logger.info("#{COMMON::syori_cnt.to_s.rjust(2)} Excelが起動しました。")

                return nil, {excel: excel, workb: workbook}
            rescue => ex
                return "method - " + __method__.to_s + " : " + ex.message + ": Excelアプリケーションの起動処理でエラーが発生しました。", {}
            end
        end

        # Excelの終了
        def proc_end(excel)
            begin
                excel.quit

                COMMON::syori_cnt += 1
                $logger.info("#{COMMON::syori_cnt.to_s.rjust(2)} Excelが終了しました。")

                # エクスプローラーを起動
                system("explorer #{$has_local["convpdf"]["path_pdf"].gsub("/", "\\")}")
                return nil
            rescue => ex
                return "method - " + __method__.to_s + " : " + ex.message + ": Excelアプリケーションの終了処理でエラーが発生しました。"
            end
        end

        # PDFファイルへの出力
        def proc_crtpdf(workbook)
            begin
                sheets = workbook.sheets
                
                # シートの一覧を取得
                ary_sheets = []
                sheets.each { |sheet| ary_sheets << sheet.Name }
                
                # 不要なシートを削除
                ary_sheets.delete_if { |sheet| sheet.include?("【原紙】") }
                
                # Excelの各シートからPDFファイルを生成
                ary_sheets.each do |sheet|
                    workbook.worksheets(sheet).Activate
                    workbook.worksheets(sheet).PrintOut(
                        ActivePrinter: 'Microsoft Print to PDF',
                        Copies: 1,
                        PrintToFile: true,
                        PrToFileName: "#{$has_local["convpdf"]["path_pdf"]}\\#{sheet}.pdf"
                    )
                end

                # PDFが作成できるまで待機
                sleep(0.4)

                COMMON::syori_cnt += 1
                $logger.info("#{COMMON::syori_cnt.to_s.rjust(2)} 「#{$has_local["convpdf"]["path_pdf"]}」に以下のPDFファイルを#{ary_sheets.size}個作成しました。")
                
                # ログにPDFファイルを１行づつ表示させるため体裁を整える
                ary_sheets_log = ary_sheets.each_with_index.map do |val, idx|
                    wrk_val = " #{COMMON::syori_cnt}-#{idx + 1} 「#{val}.pdf」"
                    new_val = idx == 0 ? "#{wrk_val}" : "#{" " * 50}#{wrk_val}"
                    idx == ary_sheets.length - 1 ? "#{new_val}" : "#{new_val}\n"
                end
                $logger.info("#{ary_sheets_log.join()}")
                
                # ドキュメントプロパティを変更
                ary_sheets.each do |sheet|
                    ret = PDFFILE.change_puropati(sheet)
                    return ret if !ret.nil?
                end
                
                COMMON::syori_cnt += 1
                $logger.info("#{COMMON::syori_cnt.to_s.rjust(2)} #{ary_sheets.size}個のPDFファイルのプロパティ値を編集しました。")
                
                return nil
            rescue => ex
                return "method - " + __method__.to_s + " : " + ex.message + ": PDFファイルへの出力処理でエラーが発生しました。"
            end
        end
    end
end

# PDF処理
class PDFFILE
    class << self
        # ドキュメントプロパティを変更
        def change_puropati(file)
            begin
                doc = HexaPDF::Document.open("#{$has_local["convpdf"]["path_pdf"]}\\#{file}.pdf")
                doc.trailer[:Info][:Title] = "#{file}"
                doc.trailer[:Info][:Author] = "日本ソフト開発株式会社"
                doc.write("#{$has_local["convpdf"]["path_pdf"]}\\#{file}.pdf", optimize: true)
                return nil
            rescue => ex
                return "method - " + __method__.to_s + " : " + ex.message + ": PDFファイルのドキュメントプロパティの変更処理でエラーが発生しました。"
            end
        end
    end
end

cat = nil

cat = catch(:goto_err) do

    # 初期値のセット
    COMMON.initialize

    # 開始処理
    ret = COMMON.proc_init
    throw :goto_err, ret if !ret.nil?

    # ファイルの存在チェック
    ret = COMMON.check_file
    throw :goto_err, ret if !ret.nil?
    
    module Excel; end

    # Excelの起動
    ret, proc = EXCEL.proc_init
    throw :goto_err, ret if !ret.nil?
    
    # PDFファイルへの出力
    ret = EXCEL.proc_crtpdf(proc[:workb])
    throw :goto_err, ret if !ret.nil?

    # Excelの終了
    ret = EXCEL.proc_end(proc[:excel])
    throw :goto_err, ret if !ret.nil?

    throw :goto_err, nil
end

# 終了処理
COMMON.proc_end(cat)