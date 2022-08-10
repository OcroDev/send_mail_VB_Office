# send_mail_VB_Office
Envío de correos electrónicos a diferentes destinatarios utilizando Microsoft Office y Outlook
Se utiliza con el sistema de correspondencia de Microsoft office, para hacer  los archivos individuales,
luego en el apartado de programación de Office Word se utiliza el código del archivo Source_code_mail, y se ejecuta
como una macro.

Para que el codigo fuente pueda detectar el nombre de quien recibe y el correo se debe colocar en la página del documento dos textos:

' _nombrearchivo/ (para indicar el nombre del archivo del destinatario)

' _correoelectronico/ (para que el programa seleccione el correo )

en la variable mensaje_correo, ya se establece un saludo (Buen dia) seguido del nombre de la persona destino, por lo que solo se debe
escribir el texto del mensaje sin realizar el saludo o modificar el valor de esta variable
