# Reports_Updater

1.Docker Build: docker build -t reports_updater_container .

2.Docker Run: docker run -it -d --name reports_updater_container reports_updater_container

3.Docker Exec: docker exec -it reports_updater_container /bin/sh

4.Configurar Crontab: 
    crontab -e
              

0 */6 * * * docker exec -i -u root  reports_updater_container /usr/local/bin/python /app/1.-BD_EventosAlumaule.py  >> /home/independencia/Escritorio/proyectos/Reports_Updater/log/BD_EventosAlumaule.log 2>&1
0 */6 * * * docker exec -i -u root  reports_updater_container /usr/local/bin/python /app/2.-BD_CargaAnticiposEEPP.py  >> /home/independencia/Escritorio/proyectos/Reports_Updater/log/BD_CargaAnticiposEEPP.log 2>&1
0 */6 * * * docker exec -i -u root  reports_updater_container /usr/local/bin/python /app/3.-BD_PostVenta.py  >> /home/independencia/Escritorio/proyectos/Reports_Updater/log/BD_PostVenta.log 2>&1
0 */6 * * * docker exec -i -u root  reports_updater_container /usr/local/bin/python /app/4.-UF_Drive.py  >> /home/independencia/Escritorio/proyectos/Reports_Updater/log/UF_Drive.log 2>&1
0 */6 * * * docker exec -i -u root  reports_updater_container /usr/local/bin/python /app/5.-BD_CargaProveedores.py  >> /home/independencia/Escritorio/proyectos/Reports_Updater/log/BD_CargaProveedores.log 2>&1
0 */6 * * * docker exec -i -u root  reports_updater_container /usr/local/bin/python /app/6.-BD_SaldosPorDocumento.py  >> /home/independencia/Escritorio/proyectos/Reports_Updater/log/BD_SaldosPorDocumento.log 2>&1


docker exec -i -u root  reports_updater_container /usr/local/bin/python /app/1.-BD_EventosAlumaule.py  >> /home/independencia/Escritorio/proyectos/Reports_Updater/log/BD_EventosAlumaule.log 2>&1
docker exec -i -u root  reports_updater_container /usr/local/bin/python /app/2.-BD_CargaAnticiposEEPP.py  >> /home/independencia/Escritorio/proyectos/Reports_Updater/log/BD_CargaAnticiposEEPP.log 2>&1
docker exec -i -u root  reports_updater_container /usr/local/bin/python /app/3.-BD_PostVenta.py  >> /home/independencia/Escritorio/proyectos/Reports_Updater/log/BD_PostVenta.log 2>&1
docker exec -i -u root  reports_updater_container /usr/local/bin/python /app/4.-UF_Drive.py  >> /home/independencia/Escritorio/proyectos/Reports_Updater/log/UF_Drive.log 2>&1
docker exec -i -u root  reports_updater_container /usr/local/bin/python /app/5.-BD_CargaProveedores.py  >> /home/independencia/Escritorio/proyectos/Reports_Updater/log/BD_CargaProveedores.log 2>&1
docker exec -i -u root  reports_updater_container /usr/local/bin/python /app/6.-BD_SaldosPorDocumento.py  >> /home/independencia/Escritorio/proyectos/Reports_Updater/log/BD_SaldosPorDocumento.log 2>&1


