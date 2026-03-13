from services import utils_download_files
from services import utils_attribute
from services import utils_import
from services import utils_excel
from services import utils_merge

class Utils:

    @staticmethod
    def procesar_excel_multimedia(ruta_excel, carpeta_salida):
        return utils_download_files.procesar_excel(ruta_excel, carpeta_salida)

    @staticmethod
    def preparar_excel_import(ruta_excel):
        return utils_excel.preparar_excel(ruta_excel)

    @staticmethod
    def seleccionar_excel():
        return utils_excel.seleccionar_excel_()

    @staticmethod
    def save_to_excel(items, filename, region):
        return utils_excel.save_to_excel_(items, filename, region)

    @staticmethod
    def excel_read_and_parse(filename, region):
        return utils_excel.excel_read_and_parse_(filename, region)

    @staticmethod
    def get_or_create_attribute(attribute_name, atributos_cache, params):
        return utils_attribute.get_or_create_attribute_(attribute_name, atributos_cache, params)

    @staticmethod
    def get_attribute(attribute_name, params):
        return utils_attribute.get_attribute_(attribute_name, params)

    @staticmethod
    def create_attribute(attribute_name, params):
        return utils_attribute.create_attribute_(attribute_name, params)

    @staticmethod
    def get_or_create_attribute_value(attribute_id, value_name, valores_cache, params):
        return utils_attribute.get_or_create_attribute_value_(attribute_id, value_name, valores_cache, params)

    @staticmethod
    def create_attribute_line(product_id, attribute_id, value_id, params):
        return utils_attribute.create_attribute_line_(product_id, attribute_id, value_id, params)

    @staticmethod
    def delete_all_attributes(params):
        return utils_attribute.delete_all_attributes_(params)

    @staticmethod
    def cargar_atributos_existentes(params):
        return utils_attribute.cargar_atributos_existentes_(params)

    @staticmethod
    def cargar_valores_atributos_existentes(params):
        return utils_attribute.cargar_valores_atributos_existentes_(params)

    @staticmethod
    def sanitize_float_field(value):
        return utils_import.sanitize_float_field_(value)

    @staticmethod
    def image_url_to_base64(url):
        return utils_import.image_url_to_base64_(url)

    @staticmethod
    def get_by_x_id_interno(model_name, old_id, db, uid, password, models):
        return utils_import.get_by_x_id_interno_(model_name, old_id, db, uid, password, models)

    @staticmethod
    def get_name_by_id(model, record_id, db, uid, password, models):
        return utils_import.get_name_by_id_(model, record_id, db, uid, password, models)

    @staticmethod
    def update_odoo_product(product_name, update_values):
        return utils_import.update_odoo_product_(product_name, update_values)

    @staticmethod
    def build_product_data_from_row(row, utils, df):
        return utils_import.build_product_data_from_row_(row, utils, df)

    @staticmethod
    def update_product_description(default_code, website_description):
        return utils_import.update_product_description_(default_code, website_description)

    @staticmethod
    def preparar_categorias_para_producto(row, product_data, params):
        return utils_import.preparar_categorias_para_producto_(row, product_data, params)

    @staticmethod
    def preparar_marca_para_producto(row, product_data, params):
        return utils_import.preparar_marca_para_producto_(row, product_data, params)

    @staticmethod
    def ico_match(nombre, valor):
        return utils_import.ico_match_(nombre, valor)

    @staticmethod
    def detectar_skus_unicos(excel_path):
        return utils_merge.detectar_skus_unicos_(excel_path)