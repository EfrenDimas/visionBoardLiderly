# VisionBoardLiderly
Proyecto en vb6, conexion con db, uso de variables de entorno. podremos crear nuevas metas y cambiar de estado cada una de ellas. cuenta con login.

## Descripción

Este es un proyecto de gestión de metas para usuarios que permite crear, actualizar, visualizar y completar metas. Los usuarios pueden ingresar sus objetivos, establecer fechas de vencimiento, cargar imágenes y seguir el progreso de sus metas. La aplicación está construida en Visual Basic 6.0 (VB6) y utiliza una base de datos SQL Server para almacenar la información de los usuarios y sus metas.
## Integrantes del Proyecto

- **Efren Reyes**
- **Moisés Reyes**
- **Montserrat Aguilar**

## Detalles Técnicos

**Lenguaje de Programación**: Visual Basic 6.0 (VB6)
**Base de Datos**: SQL Server
**Conexión a la Base de Datos**: ADO (ActiveX Data Objects)
**Entorno de Desarrollo**: VB6 y SQL Server
**Variables de Entorno**: El proyecto utiliza variables de entorno para obtener las configuraciones del servidor y la base de datos. Las variables son:
  - serverTestLiderly (Servidor de la base de datos)
  - dbTestLiderly (Nombre de la base de datos)
  
  Estas deben estar configuradas correctamente en el entorno para que la conexión a la base de datos funcione sin problemas.

  ## Pasos para Hacerlo Funcionar en Otro Equipo

1. **Instalar el Entorno de Desarrollo**:
   - Asegúrate de tener instalado **Visual Basic 6.0** en el equipo. Si no lo tienes, puedes buscarlo e instalarlo desde fuentes confiables.
   - Instalar **SQL Server** o asegurarte de tener acceso a una base de datos SQL Server configurada correctamente.  
   - 
2. **Crear las Tablas en la Base de Datos**:

El siguiente script SQL se utiliza para crear las tablas necesarias en la base de datos, basadas en las clases User y Goal del proyecto.  

```sql
-- Crear tabla de Usuarios
CREATE TABLE Users (
    userID INT PRIMARY KEY IDENTITY(1,1),
    username NVARCHAR(50) NOT NULL UNIQUE,
    Password NVARCHAR(255) NOT NULL,
    FullName NVARCHAR(100),
    Email NVARCHAR(100),
    CreatedDate DATETIME DEFAULT GETDATE(),
    LastLoginDate DATETIME
);

-- Crear tabla de Metas
CREATE TABLE Goals (
    GoalID INT PRIMARY KEY IDENTITY(1,1),
    userID INT NOT NULL,
    Title NVARCHAR(255) NOT NULL,
    Category NVARCHAR(100),
    Description NVARCHAR(500),
    DueDate DATETIME,
    ImagePath NVARCHAR(255),
    CreatedDate DATETIME DEFAULT GETDATE(),
    Status NVARCHAR(50),
    FOREIGN KEY (userID) REFERENCES Users(userID)
);

-- Insertar un usuario de ejemplo
INSERT INTO Users (username, Password, FullName, Email) 
VALUES ('admin', 'adminpassword', 'Admin User', 'admin@example.com');

-- Insertar una meta de ejemplo para el usuario admin
INSERT INTO Goals (userID, Title, Category, Description, DueDate, Status) 
VALUES (1, 'Aprender VB6', 'Desarrollo', 'Meta para aprender Visual Basic 6.0', '2024-12-31', 'En progreso');   

```
 


3. **Configurar las Variables de Entorno**:
   - Asegúrate de configurar las variables de entorno necesarias para la conexión con la base de datos:
     - serverTestLiderly: Define la dirección del servidor de base de datos.
     - dbTestLiderly: Define el nombre de la base de datos.
   - Estas variables se utilizan para establecer la conexión con la base de datos en el código.

4. **Restaurar la Base de Datos**:
   - Restaura o configura la base de datos SQL Server que contiene las tablas necesarias (Users y Goals) en el equipo. Puedes utilizar el archivo .bak o ejecutar los scripts de creación de tablas proporcionados.

5. **Abrir el Proyecto en VB6**:
   - Abre el proyecto en Visual Basic 6.0 (con el archivo .vbp o .frm correspondiente).

6. **Ejecutar la Aplicación**:
   - Ejecuta el proyecto desde VB6 y asegúrate de que la conexión con la base de datos funcione correctamente. Si la configuración es correcta, deberías poder agregar y gestionar metas.



## Capturas de Pantalla

Aquí van los espacios donde colocarás las capturas de pantalla correspondientes.

1. **Pantalla aplicacion**  
   _Captura del ejecutable del proyecto._

   ![image](https://github.com/user-attachments/assets/9119d952-3b54-4c17-bb76-e726957d8b61)


2. **Formulario de Inicio de Sesion**  
   _Captura de la interfaz para logearse._

   ![image](https://github.com/user-attachments/assets/f0e51f4a-9e5f-4877-8cac-1fe5013d9388)
3. **Pantalla de Inicio**  
   _Captura de la interfaz de inicio y bienvenida al usuario._
  ![image](https://github.com/user-attachments/assets/578f1922-7334-4dfb-8d45-dfedf59934ec)
4. **Pantalla de listado de Metas**  
   _Captura de la interfaz donde lista las metas del usuario._
  ![image](https://github.com/user-attachments/assets/e49d4331-2fc8-4f45-bcba-777789dfb03b)
5. **Pantalla de Imagen no encontrada**  
   _Captura de la interfaz donde se muestra una imagen por defecto cuando no se encuentra._
  ![image](https://github.com/user-attachments/assets/cfbd7088-e106-4446-8416-7b2ac89ca64d)
6. **Pantalla de nueva Meta**  
   _Captura de la interfaz para el registro de una nueva meta._
  ![image](https://github.com/user-attachments/assets/f8b8b322-b5b2-445a-9ade-0c4238ea2afc)
  ![image](https://github.com/user-attachments/assets/3d9bd408-810f-412b-a20c-65ab65dead0f)
  ![image](https://github.com/user-attachments/assets/2e3389a5-4b96-4755-9ac0-971f816dfc12)
  ![image](https://github.com/user-attachments/assets/ad42f261-7e1b-42b5-ba47-5144e2986fc4)
  ![image](https://github.com/user-attachments/assets/20cb3c07-9966-4461-81f0-abc066a0ff2f)


7. **Formulario de Detalles de Meta**  
   _Captura de la pantalla donde se visualiza y se completan las metas._

   ![image](https://github.com/user-attachments/assets/6d37dbea-1580-4c1e-9038-98cc3f3a0dc8)

8. **Pantalla con Validación de Status**  
   _Captura que muestra el cambio de color de un `Label` según el estado de la meta._

   ![image](https://github.com/user-attachments/assets/a643930a-6d7f-4e3b-9f51-474ead44b2d3)

9. **Pantalla de Metas Completadas**  
   _Captura que muestra el estado de las metas cuando son completadas y se oculta el boton de completar._

   ![image](https://github.com/user-attachments/assets/58bfe13f-df65-4e12-ac1e-439d02259042)
10. **Pantalla Responsive**  
   _Captura que muestra la aplicacion responsive._

  ![image](https://github.com/user-attachments/assets/59414881-f89f-446f-85c2-c4e5ab77d055)


## Propósitos de Año Nuevo y Reflexiones

### Propósito de Año Nuevo

#### **Efren**: 
- **Propósito**: Mi propósito de este año es seguir mejorando en el desarrollo de software, enfocándome especialmente en aprender nuevas tecnologías que me ayuden a crear soluciones más eficientes y escalables. También quiero mejorar mi capacidad para colaborar en equipo, compartiendo más ideas y ayudando a mis compañeros cuando lo necesiten.
- **Reflexión**: Este año he aprendido la importancia de ser organizado y de gestionar bien mi tiempo. Aunque han habido desafíos, los he visto como oportunidades para crecer. La clave está en no rendirse, sino aprender de los errores y seguir adelante.

#### **Montserrat ;)** :
- **Propósito**: Mi objetivo este año es fortalecer mis habilidades en programación de bases de datos y optimización de código. También me propongo mejorar mi capacidad de análisis y solución de problemas, enfrentando con más creatividad los desafíos técnicos que surjan.
- **Reflexión**: Este año me ha enseñado mucho sobre la importancia de la paciencia y la perseverancia. No siempre las cosas salen como esperamos, pero he aprendido a ver los obstáculos como oportunidades para innovar y encontrar nuevas soluciones.

#### **Moisés**:
- **Propósito**: Mi propósito para este año es adquirir una mayor experiencia trabajando en proyectos en equipo, además de profundizar mis conocimientos en desarrollo web y en tecnologías emergentes como la inteligencia artificial. También quiero ser más proactivo y tomar más responsabilidades en los proyectos.  
- **Reflexión**: Este año fue un desafío en términos de adaptación, pero he crecido mucho como programador. Aprendí que la comunicación es crucial dentro de un equipo, y cada miembro aporta algo único que enriquece el proyecto. Este año me siento más preparado para los retos del futuro.
