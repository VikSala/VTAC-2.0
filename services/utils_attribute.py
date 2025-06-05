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

def get_or_create_attribute_(attribute_name, params):
    """
        Verifica si el atributo existe. Si no, lo crea y devuelve su ID.
    """
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

    if not existing_attributes:
        # Si el atributo no existe, lo creamos
        attribute_vals = {'name': attribute_name}
        attribute_id = models.execute_kw(
            db, uid, password,
            'product.attribute', 'create',
            [attribute_vals]
        )
        #print(f"Atributo '{attribute_name}' creado con ID {attribute_id}")
        return attribute_id
    else:
        # Si el atributo ya existe, devolvemos su ID
        attribute_id = existing_attributes[0]['id']
        #print(f"Atributo '{attribute_name}' ya existe con ID {attribute_id}")
        return attribute_id

def get_or_create_attribute_value_(attribute_id, value_name, params):
    """
        Verifica si el valor del atributo existe. Si no, lo crea y devuelve su ID.
    """
    models = params.models
    db = params.db
    uid = params.uid
    password = params.password

    # Buscar el valor del atributo
    existing_values = models.execute_kw(
        db, uid, password,
        'product.attribute.value', 'search_read',
        [[['attribute_id', '=', attribute_id], ['name', '=', value_name]]],
        {'fields': ['id', 'name']}
    )

    if not existing_values:
        # Si el valor no existe, crear uno nuevo
        value_id = models.execute_kw(
            db, uid, password,
            'product.attribute.value', 'create',
            [{
                'name': value_name,
                'attribute_id': attribute_id,
            }]
        )
        #print(f"Valor '{value_name}' creado con ID {value_id}")
        return value_id
    else:
        value_id = existing_values[0]['id']
        #print(f"Valor '{value_name}' ya existe con ID {value_id}")
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
        print(
            f"✅ Línea de atributo actualizada para producto {product_id}, atributo {attribute_id} con valor {value_id}")
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