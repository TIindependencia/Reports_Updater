# Reports_Updater

1.Docker Build: docker build -t reports_updater_container .

2.Docker Run: docker run -it -d --name reports_updater_container reports_updater_container

3.Docker Exec: docker exec -it reports_updater_container /bin/sh

4.Configurar Crontab: 
              1.- crontab -e
              
              2.- 0 */6 * * * docker exec -i -u root  reports_updater_container /usr/local/bin/python /app/BDINFORMEENVENTOSALUMAULE.py  >> /home/independencia/Escritorio/Reports_Updater/log/Alumaule.log 2>&1






