clc;
clear;
close all;

%% Problem definition
% NC: number of components
% NP: number of phases
%
% z  : overall composition vector, size = 1 x NC
% K  : equilibrium ratio matrix, size = (NP-1) x NC

z = [0.0514653886412099, 0.138454311076800, 0.179124388193699, 0.00737965098936962, 0.161213159691314, 0.230848099529865, 0.231515001877742];

K = [30.3161162376694, 1.12470173657839, 0.289545796442771, 1.06349369764007, 0.937942923643921, 1.10339573782577, 2.39957187574342;
     32.8025838160792, 0.935147920032337, 0.210129352682663, 431.406714443393, 0.678719739868416, 0.632294134553328, 1.80735234845429];

tol = 1.0e-8;

%% Initial estimates
initialOkuno    = Initial_Estimation_Okuno(z, K);
initialHuang    = Initial_Estimation_Huang(z, K);
initialGradient = Initial_Estimation_Gradient_Huang(z, K);

fprintf('Initial estimate (Okuno):\n');
fprintf('%.6f ', initialOkuno);
fprintf('\n');

fprintf('Initial estimate (Huang):\n');
fprintf('%.6f ', initialHuang);
fprintf('\n');

fprintf('Initial estimate (Gradient):\n');
fprintf('%.6f ', initialGradient);
fprintf('\n');

%% Solve the Rachford-Rice equations
betaOkuno    = RR_Okuno(z, K, initialOkuno, tol);
betaHuang    = RR_Huang(z, K, initialHuang, tol);
betaGradient = RR_Huang(z, K, initialGradient, tol);

%% Display results
fprintf('Beta from Okuno:\n');
fprintf('%.6f ', betaOkuno);
fprintf('\n');

fprintf('Beta from Huang:\n');
fprintf('%.6f ', betaHuang);
fprintf('\n');

fprintf('Beta from Gradient:\n');
fprintf('%.6f ', betaGradient);
fprintf('\n');