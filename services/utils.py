from services import utils_download_files
from services import utils_attribute
from services import utils_import
from services import utils_excel

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
    def save_to_excel(items, filename):
        return utils_excel.save_to_excel_(items, filename)

    @staticmethod
    def excel_read_and_parse(filename):
        return utils_excel.excel_read_and_parse_(filename)

    @staticmethod
    def get_or_create_attribute(attribute_name, params):
        return utils_attribute.get_or_create_attribute_(attribute_name, params)

    @staticmethod
    def get_attribute(attribute_name, params):
        return utils_attribute.get_attribute_(attribute_name, params)

    @staticmethod
    def create_attribute(attribute_name, params):
        return utils_attribute.create_attribute_(attribute_name, params)

    @staticmethod
    def get_or_create_attribute_value(attribute_id, value_name, params):
        return utils_attribute.get_or_create_attribute_value_(attribute_id, value_name, params)

    @staticmethod
    def create_attribute_line(product_id, attribute_id, value_id, params):
        return utils_attribute.create_attribute_line_(product_id, attribute_id, value_id, params)

    @staticmethod
    def delete_all_attributes(params):
        return utils_attribute.delete_all_attributes_(params)

    @staticmethod
    def sanitize_float_field(value):
        return utils_import.sanitize_float_field_(value)

    @staticmethod
    def image_url_to_base64(url):
        return utils_import.image_url_to_base64_(url)

    @staticmethod
    def update_odoo_product(product_name, update_values):
        return utils_import.update_odoo_product_(product_name, update_values)

    @staticmethod
    def build_product_data_from_row(row, utils, df):
        return utils_import.build_product_data_from_row_(row, utils, df)

    @staticmethod
    def get_or_create_categories(row, params, product_id):
        return utils_import.get_or_create_categories_(row, params, product_id)