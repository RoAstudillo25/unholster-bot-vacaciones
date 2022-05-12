
Proyecto realizado utilizando Amazon Web Services que se complementa por los servicios de Bucket de S3, Simple Email Service (SES) y Lambda.

Para el funcionamiento de este lambda se requiere de las dependencias de "openpyxl", las que deben ser comprimidas junto con los scripts desarrollados en lenguaje Python.

Este proyecto consiste en obtener uno o varios reportes desde la plataforma de Rex+, reportes que tendrán los datos de todo aquel empleado (Unholsteriano) que se encuentre de vacaciones o con permiso administrativo. Dicha plataforma proporciona herramientas de automatización que evita realizar la descarga manual de los reportes cada vez que sean necesarios. Esta automatización es capaz de obtener tanto el reporte de vacaciones como el de permisos administrativos y enviarlos a través de un correo, que proporciona de manera automática la plataforma, hacia un destino el cual será el correo creado con el fin de ser utilizado para este poryecto, cabe destacar que esta automatización puede llegar a ser programada, enviando estos reportes una vez por semana, cada dos semanas, o el intervalo de tiempo que ser requiera para hacer uso de los reportes. El correo empresarial, que es para lo que se creó, es configurado en la herramienta SES de AWS, y permite una vez llega un correo a esta dirección de email, almacenar por completo dicho correo en un bucket que proporciona el servicio de S3 con un formato MIME. El Lambda por su parte, es configurado para ser invocado una vez se almacena un correo en el bucket, y ejecuta la función desarrollada para extraer y filtrar los datos, para finalmente ser mostrado en la plataforma de comunicación interna que ocupa Unholster, llamada Slack, a todos los miembros que participan en el canal #general para asi tener conocimiento de quien se encuentra ausente y des esta forma saber con quien contar, y con quien no, para la realización de los diversos proyectos que se trabajan día a día, y además mostrar quienes proximamente no estarán presentes ya sea por motivos de vacaciones o algun permiso  otorgado.