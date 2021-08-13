import traceback
import util
from CopyBom import CopyBom

logger = util.get_logger(__file__)

if __name__ == '__main__':
    s, config = util.load_config()

    if s:

        try:
            bom_path = config["bom_path"]
            bom_list_file = config["bom_list_file"]

            # Instantiate
            cb = CopyBom(bom_path, bom_list_file)

            # 建folder + 取得該程式的路徑
            cb.make_dir()

            # 從EXCEL中取出每一個BOM的(第一個字, BOM檔名)
            # cb.action中把它們複製到指定路徑
            for prefix, file_name in cb.parse_bom():
                cb.action(prefix, file_name)

            cb.save()

        except SystemError as e:
            logger.error("錯誤:{}".format(traceback.format_exc()))

    else:
        logger.error("沒有設定檔或設定檔異常：{}".format(traceback.format_exc()))