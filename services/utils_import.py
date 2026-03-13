from services.campos_odoo import ClavesExcel, excel_a_odoo
import pandas as pd
import ast
from deep_translator import GoogleTranslator

DESCRIPTION_HTML_1 = """<ul class="nav nav-tabs justify-content-center" role="tablist">
  <li class="nav-item o_not_editable" data-oe-model="ir.ui.view" data-oe-id="3823" data-oe-field="arch" data-oe-xpath="/t[1]/div[1]/ul[1]/li[2]">
    <a class="nav-link active" data-bs-toggle="tab" href="#tp-product-description-tab" role="tab" aria-selected="true">
      <span class="fa fa-file-text-o me-1"></span> Descripción </a>
  </li>
  <li class="nav-item o_not_editable" data-oe-model="ir.ui.view" data-oe-id="3823" data-oe-field="arch" data-oe-xpath="/t[1]/div[1]/ul[1]/li[3]">
    <a class="nav-link" data-bs-toggle="tab" href="#tp-product-specification-tab" role="tab" aria-selected="false">
      <span class="fa fa-sliders me-1"></span> Especificaciones </a>
  </li>
</ul>
<div class="tab-content">
  <div class="tab-pane fade active show" id="tp-product-description-tab" role="tabpanel">
    <div class="container-fluid">
      <div class="row m-0 py-2">
        <div class="col-12">
          <div itemprop="description" class="oe_structure" id="product_full_description" data-oe-xpath="/t[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]" data-oe-model="product.template" data-oe-id="1715" data-oe-field="website_description" data-oe-type="html" data-oe-expression="product.website_description">
            <div class="product-description" style="margin-left: 16.6667%;margin-right: 16.6667%;">
              <div></div> """
DESCRIPTION_HTML_1_END = """
            </div>
          </div>
        </div>
      </div>
    </div>
  </div>
  <div class="tab-pane fade" id="tp-product-specification-tab" role="tabpanel"></div>
</div>"""
DESCRIPTION_HTML_2 = """
 <div class="tab-pane fade" id="tp-product-downloads-tab" role="tabpanel">
  <div class="container-fluid">
    <div class="row m-0 py-2">
      <div class="col-12">
        <div class="oe_structure" id="product_pdf_downloads"
             data-oe-xpath="/t[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[2]"
             data-oe-model="product.template"
             data-oe-id="1715"
             data-oe-field="x_pdf_urls"
             data-oe-type="html"
             data-oe-expression="product.x_pdf_urls">
          <div class="row">"""
DESCRIPTION_HTML_2_END = """
          </div>
        </div>
      </div>
    </div>
  </div>
</div>
"""
SCRIPT_JS_HTML = """
<script>
  document.addEventListener("DOMContentLoaded", function() {{
    var el = document.getElementById("product_full_spec") || document.getElementById("product_attributes_simple");
    if (el) {{
      el.style.marginTop = "40px";
      el.style.marginLeft = "16.6667%";
      el.style.borderTop = "none";
    }}
    var elementToMove = document.getElementById("product_attributes_simple");
    var targetContainer = document.getElementById("tp-product-specification-tab");

    if (elementToMove && targetContainer) {{
        targetContainer.appendChild(elementToMove);
    }}
    document.getElementById("oe_structure_website_sale_product_2").innerHTML = `
        <hr>
        <div class="oe_structure oe_empty oe_structure_not_nearest mt16" id="oe_structure_website_sale_product_2" data-oe-model="ir.ui.view" data-oe-id="2195" data-oe-field="arch" data-oe-xpath="/t[1]/t[4]/div[1]/div[3]" data-editor-message="DROP BUILDING BLOCKS HERE TO MAKE THEM AVAILABLE ACROSS ALL PRODUCTS" style="
            margin-bottom: 16px !important;
        ">
          <div id="icons" class="row" style="justify-content: center;">
          </div>
        </div>`
    var iconsContainer = document.querySelector("#icons");
    var tempBlock = document.querySelector("#temp_icon_block");
    if (iconsContainer && tempBlock) {
      iconsContainer.appendChild(tempBlock);
    }
  }});
  document.addEventListener("DOMContentLoaded", function () {{
    function toggleProductFullSpec() {{
      var descriptionTab = document.querySelector('a[href="#tp-product-description-tab"]');
      var productFullSpec = document.getElementById("product_full_spec") || document.getElementById("product_attributes_simple");

      if (descriptionTab && productFullSpec) {{
        if (descriptionTab.classList.contains("active")) {{
          productFullSpec.style.display = "none";
        }} else {{
          productFullSpec.style.display = "block";
        }}
      }}
    }}

    // Ejecutar al cargar
    toggleProductFullSpec();

    // Escuchar cambios de pestaña
    var tabLinks = document.querySelectorAll('[data-bs-toggle="tab"]');
    tabLinks.forEach(function (tab) {{
      tab.addEventListener("shown.bs.tab", function () {{
        toggleProductFullSpec();
      }});
    }});
    
    function hideIfZero(valueId, boxId) {
        const el = document.getElementById(valueId);
        const box = document.getElementById(boxId);
        if (el && box) {
          const num = parseInt(el.innerText.trim()) || 0;
          if (num === 0) {
            box.style.display = "none";
          }
        }
      }

      hideIfZero("europeo", "box_europeo");
      hideIfZero("nacional", "box_nacional");
      hideIfZero("alicante", "box_alicante");
      const style = document.createElement("style");
      style.textContent = "h6.mb-1 span p { display: inline !important; }";
      document.head.appendChild(style);
  }});
</script>
"""

import base64
import requests
import time

def image_url_to_base64_(url, max_retries=3, timeout=10):
    """
    Descarga una imagen desde una URL y la convierte en base64.
    - Reintenta hasta max_retries veces si hay fallos.
    - Timeout máximo de 'timeout' segundos por petición.
    - Devuelve None si no se logra descargar.
    """

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/113.0.0.0 Safari/537.36"
        ),
        "Referer": url,
    }

    for intento in range(1, max_retries + 1):
        try:
            start = time.time()
            response = requests.get(url, headers=headers, timeout=timeout)
            elapsed = round(time.time() - start, 2)

            if response.status_code == 200:
                print(f"✅ Imagen descargada correctamente ({elapsed}s) → {url}")
                return base64.b64encode(response.content).decode("utf-8")
            else:
                print(f"⚠️ [{intento}/{max_retries}] Respuesta {response.status_code} ({elapsed}s) → {url}")

        except requests.exceptions.Timeout:
            print(f"⏳ [{intento}/{max_retries}] Timeout ({timeout}s) en {url}")
        except Exception as e:
            print(f"⚠️ [{intento}/{max_retries}] Error al descargar {url}: {e}")

        # Espera antes de reintentar (excepto en el último)
        if intento < max_retries:
            time.sleep(2)

    print(f"❌ No se pudo descargar la imagen tras {max_retries} intentos → {url}")
    return None

def get_name_by_id_(model, record_id, db, uid, password, models):

    if not record_id:
        return False

    try:
        record = models.execute_kw(
            db, uid, password,
            model, "read",
            [record_id],
            {"fields": ["name"], "limit": 1}
        )

        if record and isinstance(record, list):
            return record[0].get("name", False)

        return False

    except Exception as e:
        print(f"⚠️ Error obteniendo name de {model} id {record_id}: {e}")
        return False

def get_by_x_id_interno_(model_name, old_id, db, uid, password, models):
    """
    Busca un registro en DESTINO por el campo x_id_interno.
    :return: id del registro en DESTINO o None
    """

    ids = models.execute_kw(
        db, uid, password,
        model_name, "search",
        [[("x_id_interno", "=", old_id)]],
        {"limit": 1}
    )

    return ids[0] if ids else None

def safe_val_(val):
    # === Función auxiliar: convierte un NaN en None ===
    return None if pd.isna(val) else val

def safe_text_(text):
    return "" if pd.isna(text) else text

def sanitize_float_field_(value):
    # === Función auxiliar: convierte string a float ===
    value = safe_val_(value)
    if value == 'N/A' or value is None: return 0.0
    if isinstance(value, str):
        value = value.replace(',', '.')
        value = ''.join([c for c in value if c.isdigit() or c == '.'])

        try:
            return float(value)
        except ValueError:
            return 0.0
    return float(value)

def update_odoo_product_(product_name, update_values):
    """
    Actualiza un product.template en Odoo por nombre.
    """
    from core.App_Connection import models, db, uid, password
    # Buscar producto por nombre
    product_ids = models.execute_kw(
        db, uid, password,
        'product.template', 'search',
        [[['name', '=', product_name]]]
    )

    if product_ids:
        # Actualizar el primer producto encontrado
        success = models.execute_kw(
            db, uid, password,
            'product.template', 'write',
            [[product_ids[0]], update_values]
        )
        return success
    else:
        print(f"Producto '{product_name}' no encontrado.")
        return False

#Datos de Producto
def build_product_data_from_row_(fila, utils, df):
    """
    Construye el diccionario de datos de producto para crear en Odoo.
    Solo incluye campos con valor válido.
    """
    product_data = {}
    ODOO16 = False #True
    ADVANCE = False
    LEDSPRESS = False

    # Declarar campos a evitar (se importan de forma especial)
    campos_prohibidos = {
        excel_a_odoo.get(ClavesExcel.ATRIBUTOS.value),
        excel_a_odoo.get(ClavesExcel.GALERIA.value),
        excel_a_odoo.get(ClavesExcel.VIDEOS.value),
        excel_a_odoo.get(ClavesExcel.DOCUMENTOS.value),
        excel_a_odoo.get(ClavesExcel.CATEGORIA.value),
        excel_a_odoo.get(ClavesExcel.MARCA.value),
        "Iconos", "Descuento", "Resumen", "Modelo", "Logo"
    }

    for columna_excel, campo_odoo in excel_a_odoo.items():
        valor = fila.get(columna_excel)

        if campo_odoo in campos_prohibidos:
            continue

        if ADVANCE and (campo_odoo == "name" or campo_odoo == "website_description" or campo_odoo == "public_categ_ids"):
            valor = str(valor)
            if campo_odoo == "website_description" and valor:
                valor = valor.replace("_x000D_", "")
            valor = GoogleTranslator(source='pt', target='es').translate(valor)
            product_data[campo_odoo] = valor

        if campo_odoo == excel_a_odoo.get(ClavesExcel.DESCRIPCION_WEB.value):
            downloads = DESCRIPTION_HTML_2
            pdf_html = """<div class="row">"""

            documentos = fila.get(ClavesExcel.DOCUMENTOS.value)

            if documentos and isinstance(documentos, str):
                try:
                    pdf_dict = ast.literal_eval(documentos)
                    for i, (pdf_name, pdf_url) in enumerate(pdf_dict.items(), start=1):
                        bloque = f"""
                        <div class="col-md-4 mb-3">
                          <div class="d-flex justify-content-between align-items-center border rounded p-2">
                            <span class="text-truncate me-2" style="max-width: 70%;">{pdf_name}</span>
                            <a target="_blank" class="btn btn-sm btn-outline-secondary" href="{pdf_url}" title="Descargar PDF">
                              <svg width="20" height="20" viewBox="0 0 20 20" xmlns="http://www.w3.org/2000/svg">
                                <polyline fill="none" stroke="#000" points="14,10 9.5,14.5 5,10"></polyline>
                                <rect x="3" y="17" width="13" height="1"></rect>
                                <line fill="none" stroke="#000" x1="9.5" y1="13.91" x2="9.5" y2="3"></line>
                              </svg>
                            </a>
                          </div>
                        </div>
                        """
                        pdf_html += bloque
                except Exception as e:
                    print(f"⚠️ Error procesando documentos PDF: {e}")

            """ LEDSPRESS
            <div class="product-details" style="margin-bottom: 5%;">
                  <div class="sku-line">
                    SKU {sku}{modelo}
                  </div>
                  {resumen}
            </div>"""

            downloads += pdf_html + DESCRIPTION_HTML_2_END
            description = DESCRIPTION_HTML_1 + f"<p>{safe_text_(valor)}</p>{pdf_html + '</div>'}" + DESCRIPTION_HTML_1_END
            website_description = safe_text_(valor) if ODOO16 else description + SCRIPT_JS_HTML
            product_data[campo_odoo] = website_description

        elif pd.notna(valor):
            if campo_odoo in [
                excel_a_odoo.get(ClavesExcel.PRECIO.value),
                excel_a_odoo.get(ClavesExcel.COSTE.value),
                excel_a_odoo.get(ClavesExcel.PESO.value),
                excel_a_odoo.get(ClavesExcel.VOLUMEN.value)
            ]:
                valor = utils.sanitize_float_field(valor)
            elif campo_odoo == excel_a_odoo.get(ClavesExcel.IMAGEN.value):
                #if "http" not in valor: valor = "http://143.47.53.74:8070/" + valor
                valor = utils.image_url_to_base64(valor)
            elif campo_odoo == excel_a_odoo.get(ClavesExcel.REFERENCIA.value):
                valor = str(valor).strip()
                if valor.endswith(".0"):
                    valor = valor[:-2]
            else:
                valor = str(valor)

            product_data[campo_odoo] = valor

    if ODOO16: product_data["detailed_type"] = "product"
    else: product_data["is_storable"] = 1

    product_data["invoice_policy"] = "delivery"
    product_data["show_availability"] = True
    product_data["available_threshold"] = 100.000

    return product_data

def update_product_description_(default_code, website_description):
    downloads = DESCRIPTION_HTML_2
    sku = default_code
    bloque_resumen = f"""
                    <div class="product-details" style="margin-bottom: 5%;">
                          <div class="sku-line">
                            SKU {sku}
                          </div>
                    </div>"""

    SCRIPT_JS_HTML_2 = f"""
    <script>
        document.addEventListener("DOMContentLoaded", function() {{
            const target = document.querySelector('div[data-oe-field="description_ecommerce"]');
            if (target) {{
                const resumenHTML = `{bloque_resumen}`;
                // Insertar al inicio del contenedor
                target.insertAdjacentHTML('afterbegin', resumenHTML);
            }} else {{
                console.warn("⚠️ No se encontró el div data-oe-field='description_ecommerce'");
            }}
        }})
    </script>
                """

    downloads += DESCRIPTION_HTML_2_END
    description = DESCRIPTION_HTML_1 + f"<p>{website_description}</p></div>" + DESCRIPTION_HTML_1_END
    website_description = f"""
    <div>
    {description}
    {SCRIPT_JS_HTML}
    {SCRIPT_JS_HTML_2}
    </div>
    """
    return website_description

#Categoria y Marca
def _get_or_create_by_name(model_name, name_field, name_value, params):
    models, db, uid, password = params
    results = models.execute_kw(db, uid, password,
                                model_name, 'search_read',
                                [[(name_field, '=', name_value)]], {'fields': ['id'], 'limit': 1})
    if results:
        return results[0]['id']
    else:
        return models.execute_kw(db, uid, password,
                                 model_name, 'create', [{name_field: name_value}])

def _get_or_create_category_hierarchy(model_name, ruta_completa, params):
    """
    Crea jerarquía 'PADRE / HIJO / ...' sin duplicar categorías.
    Primero busca o crea el padre, luego los hijos, con validación precisa (name + parent_id).
    """
    models, db, uid, password = params
    ruta = [r.strip() for r in ruta_completa.split(" / ")]
    if not ruta:
        return None

    # Paso 1: tratar el padre
    parent_name = ruta[0]
    domain = [('name', '=', parent_name), ('parent_id', '=', False)]
    result = models.execute_kw(db, uid, password,
                               model_name, 'search_read',
                               [domain], {'fields': ['id'], 'limit': 1})

    if result:
        parent_id = result[0]['id']
    else:
        parent_id = models.execute_kw(db, uid, password,
                                      model_name, 'create', [{'name': parent_name}])

    # Paso 2: tratar los hijos (si los hay)
    for name in ruta[1:]:
        domain = [('name', '=', name), ('parent_id', '=', parent_id)]
        result = models.execute_kw(db, uid, password,
                                   model_name, 'search_read',
                                   [domain], {'fields': ['id'], 'limit': 1})

        if result:
            parent_id = result[0]['id']
        else:
            values = {'name': name, 'parent_id': parent_id}
            parent_id = models.execute_kw(db, uid, password,
                                          model_name, 'create', [values])

    return parent_id

def preparar_categorias_para_producto_(row, product_data, params):
    """
    Asigna categorías interna y públicas al diccionario `product_data` para usar en un create de Odoo.
    """
    public_categ_ids_list = []
    row_value = row[ClavesExcel.CATEGORIA.value].split("/")

    if row_value:
        rutas_categorias = row[ClavesExcel.CATEGORIA.value]

        # Asegurarse de que es lista
        if isinstance(rutas_categorias, str):
            rutas_categorias = [rutas_categorias]

        for categ_name in rutas_categorias:

            # 1. Asignar categoría interna desde la primera parte de la primera ruta
            categoria_padre = categ_name.split(" / ")[0].strip()
            internal_categ_id = _get_or_create_by_name('product.category', 'name', categoria_padre, params)
            product_data['categ_id'] = internal_categ_id

            # 2. Asignar categorías públicas completas por jerarquía
            for cats in categ_name.split(','):
                public_categ_id = _get_or_create_category_hierarchy('product.public.category', cats, params)
                if public_categ_id:
                    public_categ_ids_list.append(public_categ_id)

        # Añadir como m2m para el create
        if public_categ_ids_list:
            product_data['public_categ_ids'] = [(6, 0, public_categ_ids_list)]
        else:
            product_data.pop('public_categ_ids', None)

def preparar_marca_para_producto_(row, product_data, params):
    if row[ClavesExcel.MARCA.value]:
        product_data['product_brand_id'] = _get_or_create_by_name('product.brand', 'name',
                                                                 row[ClavesExcel.MARCA.value], params)#row[ClavesExcel.MARCA.value], params) "V-Tac"

#Iconos
def ico_match_(nombre, valor):
    ico_url = ""
    ruta = "http://79.72.55.217:8069"
    if not valor:
        if "AMAZON" in nombre or "GOOGLE" in nombre or "ALEXA" in nombre: ico_url = ruta + "/web/image/978-39bca721/DUAL%20CONTROL%20APP%20%26%20REMOTE%201.png,"
        if "BRIDGELUX" in nombre: ico_url += ruta + "/web/image/966-503077fa/BRIDGELUX%20CHIP.png,"
        if "RECARGABLE" in nombre: ico_url += ruta + "/web/image/965-edb6193b/BATTERY%20RECHARGEABLE%201.png,"
        if "insecto" in nombre: ico_url += ruta + "/web/image/981-607f50f8/INSECT.png,"
        if "INVENTRONICS" in nombre: ico_url += ruta + "/web/image/982-fc4e0017/INVENTRONICS%20DRIVER.png,"
        if "LIFUD" in nombre: ico_url += ruta + "/web/image/991-27a88b03/LIFUD%20DRIVER.png,"
        if "MEAN WELL" in nombre: ico_url += ruta + "/web/image/1004-39960077/MW%20MEAN%20WELL%202.png,"
        if "SAMSUNG" in nombre: ico_url += ruta + "/web/image/1005-f6dcb47f/SAMSUNG%20LED%20ILLUMINATED.png,"
        if "CREE" in nombre: ico_url += ruta + "/web/image/400031-f48880d5/CREE%20LED%20POWERED%202.png,"
    else:
        match nombre:
            case "Tono de luz":
                if valor == "3 en 1": ico_url = ruta + "/web/image/967-f42b7cd5/CCT%20SWITCHABLE.png"
            case "CRI":
                if valor == ">80": ico_url = ruta + "/web/image/970-e1b3be79/CRI%2080.png"
                if valor == ">90": ico_url = ruta + "/web/image/971-348cdd7c/CRI%2090.png"
                if valor == ">95": ico_url = ruta + "/web/image/972-1d46a121/CRI%2095.png"
            case "Tensión":
                if valor == "5V": ico_url = ruta + "/web/image/973-c966f225/DC%205V.png"
                elif valor == "12V": ico_url = ruta + "/web/image/974-0eb4bbf0/DC%2012V.png"
                elif valor == "24V": ico_url = ruta + "/web/image/975-4afad8e0/DC%2024V.png"
                elif valor == "48V": ico_url = ruta + "/web/image/976-7c159814/DC%2048V.png"
            case "Regulable":
                if valor == "Sí" or valor == "Si": ico_url = ruta + "/web/image/976-7c159814/DC%2048V.png"
            case "Eficacia luminosa (lm/W)":
                if valor == "100 lm/W": ico_url = ruta + "/web/image/992-3a2c91f7/LUMEN-W%20100.png"
                elif valor == "110 lm/W": ico_url = ruta + "/web/image/993-d674d229/LUMEN-W%20110.png"
                elif valor == "120 lm/W": ico_url = ruta + "/web/image/994-97f69ce1/LUMEN-W%20120.png"
                elif valor == "130 lm/W": ico_url = ruta + "/web/image/995-40aac088/LUMEN-W%20130.png"
                elif valor == "135 lm/W": ico_url = ruta + "/web/image/996-c4d19d0d/LUMEN-W%20135.png"
                elif valor == "140 lm/W": ico_url = ruta + "/web/image/997-76028f6a/LUMEN-W%20140.png"
                elif valor == "150 lm/W": ico_url = ruta + "/web/image/998-f9089c11/LUMEN-W%20150.png"
                elif valor == "160 lm/W": ico_url = ruta + "/web/image/999-2e5cc3a5/LUMEN-W%20160.png"
                elif valor == "180 lm/W": ico_url = ruta + "/web/image/1000-bbaf07e7/LUMEN-W%20180.png"
                elif valor == "185 lm/W": ico_url = ruta + "/web/image/1001-0e26b857/LUMEN-W%20185.png"
                elif valor == "200 lm/W": ico_url = ruta + "/web/image/1002-99966c1a/LUMEN-W%20200.png"
                elif valor == "210 lm/W": ico_url = ruta + "/web/image/1003-1db5523d/LUMEN-W%20210.png"
            case "Grado de protección IK":
                if valor == "IK07": ico_url = ruta + "/web/image/979-f9d3bc09/IK07%20RATING.png"
                elif valor == "IK08": ico_url = ruta + "/web/image/980-d7edf4b7/IK08%20RATING.png"
            case "Tiempo de carga":
                if valor == "12 horas": ico_url = ruta + "/web/image/968-818cfea0/CHARGING%20TIME%2012%20hr.png"
                elif valor == "24 horas": ico_url = ruta + "/web/image/969-b3139376/CHARGING%20TIME%2024%20hr.png"
            case "Grado de protección IP":
                if valor == "IP20": ico_url = ruta + "/web/image/983-87ea7948/IP20%20RATING.png"
                elif valor == "IP44": ico_url = ruta + "/web/image/984-1b0beb06/IP44%20RATING.png"
                elif valor == "IP45": ico_url = ruta + "/web/image/985-7c9a7e77/IP45%20RATING.png"
                elif valor == "IP54": ico_url = ruta + "/web/image/986-50103ff8/IP54%20RATING.png"
                elif valor == "IP65": ico_url = ruta + "/web/image/987-189a1e64/IP65%20RATING.png"
                elif valor == "IP66": ico_url = ruta + "/web/image/988-053cd661/IP66%20RATING.png"
                elif valor == "IP67": ico_url = ruta + "/web/image/989-3ee06319/IP67%20RATING.png"
                elif valor == "IP68": ico_url = ruta + "/web/image/990-0b429b8d/IP68%20RATING.png"
            #case "Rango de detección": if valor == "360º": ico_url = "LIGHT 360º.png"#???
            case "Garantía":
                if valor == "1 año": ico_url = ruta + "/web/image/1006-9570e6ea/YEAR%20WARRANTY%2001.png"
                elif valor == "2 años": ico_url = ruta + "/web/image/1007-c120527b/YEAR%20WARRANTY%2002.png"
                elif valor == "3 años": ico_url = ruta + "/web/image/1008-386c18d0/YEAR%20WARRANTY%2003.png"
                elif valor == "4 años": ico_url = ruta + "/web/image/1009-78032617/YEAR%20WARRANTY%2004.png"
                elif valor == "5 años": ico_url = ruta + "/web/image/1010-7a888e5d/YEAR%20WARRANTY%2005.png"

    return False if ico_url == "" else ico_url
