# Control del quiosco

Aplicación web para llevar productos, stock, ventas, gastos, **cierre de recreo**, resumen, historial de precios y un **catálogo con pedidos por WhatsApp**.

Todo está en un único archivo: `index.html`. No hay que instalar nada.

## 1. Subirla a GitHub

1. Entrá a https://github.com/new y creá un repositorio (por ejemplo `control-quiosco`).
2. Dentro del repositorio tocá **Add file → Upload files** y subí `index.html` y `README.md`.
   El archivo tiene que llamarse exactamente `index.html` y estar en la raíz del repositorio.
3. Tocá **Commit changes**.

## 2. Publicarla como página web (gratis)

1. En el repositorio: **Settings → Pages**.
2. En **Build and deployment → Source** elegí **Deploy from a branch**.
3. En **Branch** elegí `main` y la carpeta `/ (root)`, y tocá **Save**.
4. Esperá uno o dos minutos. Tu página queda en `https://TU_USUARIO.github.io/control-quiosco/`.

## 3. Primeros pasos

1. Entrá a **Modo administrador**. La primera vez te pide definir una clave.
2. Cargá tus productos en **Productos** (nombre, categoría, precio, costo, stock).
3. En **Catálogo** cargá tu número de WhatsApp (con código de país, sin espacios ni signos, por ejemplo `5493704123456`) y las fotos de los productos.
4. Los clientes, sin clave, solo ven el catálogo: arman su pedido con cantidades y lo envían ya armado a tu WhatsApp.
5. Después de cada recreo usá la pestaña **Cierre**: anotás lo que quedó de cada producto y la app calcula las unidades vendidas.

## 4. Dónde se guardan los datos (importante)

- **Sin configurar nada**, los datos se guardan en el navegador de *cada dispositivo*. Sirve para uso personal, pero **un cliente que abra el link desde su celular vería el catálogo vacío**, porque tus productos están en tu dispositivo.
- **Para que todos vean los mismos datos** (catálogo para clientes, o vos y un empleado), conectá una base de datos gratuita en [Supabase](https://supabase.com):

  1. Creá un proyecto en Supabase.
  2. En **SQL Editor** ejecutá:

```sql
create table kv (
  key text primary key,
  value text not null
);
alter table kv enable row level security;
create policy "lectura" on kv for select using (true);
create policy "insercion" on kv for insert with check (true);
create policy "actualizacion" on kv for update using (true) with check (true);
create policy "borrado" on kv for delete using (true);
```

  3. En **Project Settings → API** copiá la *Project URL* y la clave *anon public*.
  4. En GitHub abrí `index.html`, tocá el lápiz (Edit) y buscá `KIOSCO_CONFIG` (está cerca del principio). Pegá los dos valores entre las comillas:

```html
<script>
  window.KIOSCO_CONFIG = {
    SUPABASE_URL: "https://xxxx.supabase.co",
    SUPABASE_ANON_KEY: "tu-clave-anon",
    RECOVERY_CODE: "RECREO-4827-QX"
  };
</script>
```

  5. Tocá **Commit changes**. En un minuto la página ya usa la base compartida.

  Las claves que ya cargaste en un dispositivo no se pasan solas: al activar Supabase empezás con datos vacíos y tenés que volver a cargar los productos.

**Limitación de seguridad:** con estas reglas simples, la clave de administrador protege la pantalla pero no la base de datos. Alguien con conocimientos técnicos podría editar los datos directamente. Alcanza para empezar; el paso siguiente recomendado es agregar un inicio de sesión real (Supabase Auth) para que solo vos puedas escribir.

## 5. Si olvidás la clave de administrador

En **Modo administrador** tocá **Olvidé mi clave** e ingresá el código de recuperación. Está en `index.html`, dentro de `KIOSCO_CONFIG`, como `RECOVERY_CODE`. **Cambialo por uno propio y guardalo.**

## 6. Qué no incluye esta versión

El agente de IA (cargar por texto), el dictado por voz y la lectura de fotos de facturas funcionaban dentro de Claude.ai y **están desactivados** acá. En una web propia necesitan una clave de API de Anthropic guardada en un servidor propio (nunca dentro de la página). El resto funciona igual, incluido el Cierre de recreo, que resuelve la carga de unidades sin IA.

## 7. Si querés modificar el código

El ZIP `control-quiosco-github.zip` trae el proyecto fuente (Vite + React + Tailwind) con publicación automática en GitHub Pages. Para generar de nuevo el archivo único: `npm install` y `npm run build:single` (queda en `dist-single/index.html`).
