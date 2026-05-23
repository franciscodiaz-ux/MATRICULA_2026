/**
 * Construye el menú principal del sistema.
 * Este archivo solo contiene la definición del menú.
 */
function onOpen() {
  crearMenuPrincipal();
}

/**
 * Crea el menú principal usando una estructura declarativa.
 */
function crearMenuPrincipal() {

  const ui = SpreadsheetApp.getUi();

  // =====================================================
  // DEFINICIÓN DECLARATIVA DEL MENÚ
  // =====================================================
  const MENU = {
    nombre: "📋 Menú del Sistema",

    accionesDirectas: [
      {
        texto: "▶ Acción directa 1",
        funcion: "accionDirecta1"
      },
      {
        texto: "▶ Acción directa 2",
        funcion: "accionDirecta2"
      }
    ],

    submenus: [

      {
        nombre: "📁 Submenú 1",
        items: [
          { texto: "Función 1.1", funcion: "funcion_1_1" },
          { texto: "Función 1.2", funcion: "funcion_1_2" },
          { texto: "Función 1.3", funcion: "funcion_1_3" },
          { texto: "Función 1.4", funcion: "funcion_1_4" }
        ]
      },

      {
        nombre: "⚙️ Submenú 2",
        items: [
          { texto: "Función 2.1", funcion: "funcion_2_1" },
          { texto: "Función 2.2", funcion: "funcion_2_2" },
          { texto: "Función 2.3", funcion: "funcion_2_3" },
          { texto: "Función 2.4", funcion: "funcion_2_4" }
        ]
      },

      {
        nombre: "📊 Submenú 3",
        items: [
          { texto: "Función 3.1", funcion: "funcion_3_1" },
          { texto: "Función 3.2", funcion: "funcion_3_2" },
          { texto: "Función 3.3", funcion: "funcion_3_3" },
          { texto: "Función 3.4", funcion: "funcion_3_4" }
        ]
      }

    ]
  };

  // =====================================================
  // CREACIÓN DEL MENÚ
  // =====================================================
  const menuPrincipal = ui.createMenu(MENU.nombre);

  // Acciones directas
  MENU.accionesDirectas.forEach(item => {
    menuPrincipal.addItem(item.texto, item.funcion);
  });

  menuPrincipal.addSeparator();

  // Submenús
  MENU.submenus.forEach(submenuConfig => {

    const submenu = ui.createMenu(submenuConfig.nombre);

    submenuConfig.items.forEach(item => {
      submenu.addItem(item.texto, item.funcion);
    });

    menuPrincipal.addSubMenu(submenu);

  });

  menuPrincipal.addToUi();
}
