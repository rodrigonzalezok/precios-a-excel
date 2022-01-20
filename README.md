# precios-a-excel
Script que extrae todos los precios de los instrumentos financieros del mercado argentino y los envía a un excel en tiempo real

En este código simplemente deben colocar las rutas de: el excel a donde van a mandar los datos y en dónde tienen el driver para Selenium (recomiendo tenerlo en la misma carpeta par que la ruta sea más sencilla).

Una vez colocados las rutas, deben colocar los datos de acceso a la cuenta del broker (en este caso BullMarket) y correr el script.

El código enviará los datos en tiempo real al excel seleccionado en la ruta. 

<b>Nota importante:</b> el código tirará error siempre que esté corriendo y el excel esté abierto. Recomiendo usar ese excel como 'data base' y usar otro archivo excel para llamar los datos que necesites procesar desde tu nueva base de datos.

Espero que te sea útil, saludos!
