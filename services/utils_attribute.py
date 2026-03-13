
#region Atributos Import
def cargar_atributos_existentes_(params):
    """
    Carga todos los atributos de producto existentes en Odoo.
    Devuelve un diccionario: nombre → id
    """
    models, db, uid, password = params

    atributos = models.execute_kw(
        db, uid, password,
        'product.attribute', 'search_read',
        [[]],
        {'fields': ['id', 'name']}
    )

    return {a['name']: a['id'] for a in atributos}

def cargar_valores_atributos_existentes_(params):
    """
    Carga todos los valores de atributo existentes.
    Devuelve un diccionario: (attribute_id, value_name) → id
    """
    models, db, uid, password = params

    valores = models.execute_kw(
        db, uid, password,
        'product.attribute.value', 'search_read',
        [[]],
        {'fields': ['id', 'name', 'attribute_id'], 'limit': 100000}
    )

    cache = {}
    for v in valores:
        attr_id = v['attribute_id'][0] if isinstance(v['attribute_id'], list) else v['attribute_id']
        cache[(attr_id, v['name'])] = v['id']
    return cache

def create_attribute_(attribute_name, params):

    models = params.models
    db = params.db
    uid = params.uid
    password = params.password

    # Buscar si el atributo ya existe
    existing_attributes = models.execute_kw(
        db, uid, password,
        'product.attribute', 'search_read',
        [[['name', '=', attribute_name]]],
        {'fields': ['id', 'name']}
    )

    # Si el atributo no existe, lo creamos
    attribute_vals = {'name': attribute_name}
    attribute_id = models.execute_kw(
        db, uid, password,
        'product.attribute', 'create',
        [attribute_vals]
    )
    return attribute_id

def get_attribute_(attribute_name, params):

    models = params.models
    db = params.db
    uid = params.uid
    password = params.password

    # Buscar si el atributo ya existe
    existing_attributes = models.execute_kw(
        db, uid, password,
        'product.attribute', 'search_read',
        [[['name', '=', attribute_name]]],
        {'fields': ['id', 'name']}
    )

    attribute_id = existing_attributes[0]['id']
    return attribute_id

def get_or_create_attribute_(attribute_name, atributos_cache, params):
    """
    Verifica si el atributo existe en el diccionario. Si no, lo crea y actualiza el diccionario.
    """
    models, db, uid, password = params
    attribute_name = attribute_name.strip()

    if attribute_name in atributos_cache:
        return atributos_cache[attribute_name]

    # Si no existe, lo creamos en Odoo
    attribute_id = models.execute_kw(
        db, uid, password,
        'product.attribute', 'create',
        [{'name': attribute_name}]
    )

    atributos_cache[attribute_name] = attribute_id
    return attribute_id

def get_or_create_attribute_value_(attribute_id, value_name, valores_cache, params):
    """
    Devuelve el ID del valor del atributo. Si no existe, lo crea y actualiza el cache.
    """
    models = params.models
    db = params.db
    uid = params.uid
    password = params.password

    value_name = str(value_name)
    key = (attribute_id, value_name.strip())

    if key in valores_cache:
        return valores_cache[key]

    # Crear si no existe
    value_id = models.execute_kw(
        db, uid, password,
        'product.attribute.value', 'create',
        [{
            'name': value_name.strip(),
            'attribute_id': attribute_id,
        }]
    )

    valores_cache[key] = value_id
    return value_id

def create_attribute_line_(product_id, attribute_id, value_id, params):
    """
        Función para asignar atributo y valor al producto.
    """
    models = params.models
    db = params.db
    uid = params.uid
    password = params.password

    # Buscar si ya existe una línea de atributo para ese producto y atributo
    existing_lines = models.execute_kw(
        db, uid, password,
        'product.template.attribute.line', 'search_read',
        [[
            ('product_tmpl_id', '=', product_id),
            ('attribute_id', '=', attribute_id)
        ]],
        {'fields': ['id', 'value_ids']}
    )

    if existing_lines:
        # Ya existe, actualizar la línea
        line_id = existing_lines[0]['id']
        models.execute_kw(
            db, uid, password,
            'product.template.attribute.line', 'write',
            [[line_id], {
                'value_ids': [(6, 0, [value_id])]  # Reemplaza todos los valores anteriores con este
            }]
        )
        #print(f"✅ Línea de atributo actualizada para producto {product_id}, atributo {attribute_id} con valor {value_id}")
    else:
        # No existe, la creamos
        line_id = models.execute_kw(
            db, uid, password,
            'product.template.attribute.line', 'create',
            [{
                'product_tmpl_id': product_id,
                'attribute_id': attribute_id,
                'value_ids': [(4, value_id)]
            }]
        )

def get_odoo_field_names(params):
    # === Función auxiliar: coge los atributos disponibles de Odoo ===
    fields = params.models.execute_kw(
        params.db, params.uid, params.password,
        'product.template', 'fields_get',
        [],
        {'attributes': ['string', 'type']}
    )

    for field_name in fields:
        print(f"{field_name} → {fields[field_name]['string']}")

def delete_all_attributes_(params):
    """
    Elimina todos los atributos de producto en Odoo.
    :param params: Tupla con conexión (models, db, uid, password)
    """
    models, db, uid, password = params

    # Buscar todos los IDs de atributos
    attribute_ids = models.execute_kw(db, uid, password,
                                      'product.attribute', 'search',
                                      [[]])  # sin filtro = todos

    print(f"🔴 Se encontraron {len(attribute_ids)} atributos para borrar.")

    if not attribute_ids:
        print("✅ No hay atributos para eliminar.")
        return

    # Eliminar los atributos
    result = models.execute_kw(db, uid, password,
                               'product.attribute', 'unlink',
                               [attribute_ids])

    if result:
        print("✅ Todos los atributos fueron eliminados correctamente.")
    else:
        print("❌ Error al intentar eliminar los atributos.")

#endregion