pipeline {
    agent any

    tools {
        jdk 'jdk-12'
    }

    options {
        timestamps()
        skipStagesAfterUnstable()
        buildDiscarder(logRotator(numToKeepStr: '30'))
    }

    stages {

        stage('Clean') {
            // Only clean when the last build failed
            when {
                expression {
                    currentBuild.previousBuild?.currentResult == 'FAILURE'
                }
            }
            steps {
                sh "./gradlew clean"
            }
        }

        stage('Compile') {
            steps {
                sh './gradlew -v' // Output gradle version for verification checks
                sh "./gradlew jenkinsClean compile"
            }
        }

        stage('Integration Test') {
            steps {
                script {
                    def outputTestFolder = uk.ac.ox.ndm.jenkins.Utils.generateRandomTestFolder()
                    def port = uk.ac.ox.ndm.jenkins.Utils.findFreeTcpPort()

                    sh "./gradlew " +
                       "-Dhibernate.search.default.indexBase=${outputTestFolder} " +
                       "-Dserver.port=${port} " +
                       "integrationTest"
                }
            }
            post {
                always {
                    publishHTML([
                            allowMissing         : false,
                            alwaysLinkToLastBuild: true,
                            keepAll              : true,
                            reportDir            : 'build/reports/tests/integrationTest',
                            reportFiles          : 'index.html',
                            reportName           : 'Integration Test Report',
                            reportTitles         : 'Test'
                    ])
                    junit allowEmptyResults: true, testResults: '**/build/test-results/**/*.xml'
                    outputTestResults()
                }
            }
        }

        stage('Jacoco Report') {
            steps {
                sh "./gradlew jacocoTestReport"
            }
            post {
                always {
                    jacoco execPattern: '**/build/jacoco/*.exec'
                    publishHTML([
                            allowMissing         : false,
                            alwaysLinkToLastBuild: true,
                            keepAll              : true,
                            reportDir            : 'build/reports/jacoco/test/html',
                            reportFiles          : 'index.html',
                            reportName           : 'Coverage Report (Gradle)',
                            reportTitles         : 'Jacoco Coverage'
                    ])
                }
            }
        }

        stage('License Header Check'){
            steps{
                sh './gradlew license'
            }
        }

        stage('Deploy to Artifactory') {
            when {
                allOf {
                    anyOf {
                        branch 'master'
                        branch 'develop'
                    }
                    expression {
                        currentBuild.currentResult == 'SUCCESS'
                    }
                }

            }
            steps {
                script {
                    sh "./gradlew artifactoryPublish"
                }
            }
        }
    }

    post {
        always {
            slackNotification()
        }
    }
}
