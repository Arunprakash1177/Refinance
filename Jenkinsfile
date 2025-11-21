pipeline {

    agent any

    environment {
        DOCKER_USERNAME = "arunprakash1177"
        DOCKER_PASSWORD = credentials('dockerhub-password')
    }

    stages {

        stage('Checkout Code') {
            steps {
                git branch: 'master',
                    credentialsId: 'github-credentials',
                    url: 'https://github.com/Arunprakash1177/Refinance.git'
            }
        }

        stage('Build Docker Image') {
            steps {
                sh 'docker build -t ${DOCKER_USERNAME}/refinance .'
            }
        }

        stage('Docker Login') {
            steps {
                sh "echo ${DOCKER_PASSWORD} | docker login -u ${DOCKER_USERNAME} --password-stdin"
            }
        }

        stage('Push Image to Docker Hub') {
            steps {
                sh 'docker push ${DOCKER_USERNAME}/refinance'
            }
        }

        stage('Deploy Container') {
            steps {
                sh '''
                    # stop and remove old container if exists
                    docker rm -f refinance || true
                    
                    # pull latest image
                    docker pull arunprakash1177/refinance
                    
                    # run new container
                    docker run -d \
                        --name refinance \
                        -p 5000:5000 \
                        --restart=always \
                        arunprakash1177/refinance
                '''
            }
        }
    }
}
