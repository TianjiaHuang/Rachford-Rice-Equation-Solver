function [beta, iter, residual, point] = RR_Huang(z, K, beta0, tol, maxIter)
%RACHFORD_RICE_HUANG Solve the Rachford-Rice equations
%using a reduced feasible solution window.
%
% Reference:
%   Huang, Johns, and Birol, SPE Journal, 2026.
%
% Inputs:
%   z       : 1 x NC vector of overall compositions
%   K       : (NP-1) x NC matrix of K-values
%   beta0   : (NP-1) x 1 initial guess for phase fractions
%   tol     : convergence tolerance
%   maxIter : maximum number of Newton iterations
%
% Outputs:
%   beta     : (NP-1) x 1 vector of phase fractions
%   iter     : number of iterations performed
%   residual : infinity norm of the gradient at each iteration
%   point    : iteration history of beta values

    if nargin < 4
        tol = 1e-8;
    end
    if nargin < 5
        maxIter = 20;
    end

    % Problem size
    [NPm1, NC] = size(K); % NPm1 = number of unknown beta values

    % Line search parameters
    maxLineSearchIter = 10;
    tolLineSearch = 1e-3;

    % Preallocate iteration history
    residual = zeros(1, maxIter);
    point = zeros(maxIter, NPm1);

    % Initial guess
    beta = beta0;

    % Precompute coefficients used in the denominator:
    a = 1 - K;

    % Compute upper bounds defining the smaller feasible solution window
    % b(i) corresponds to the upper bound for component i
    theta = ones(1, NC);

    for i = 1:NC
        for j = 1:NPm1
            Kmax = max(K(j, :));
            Kmin = min(K(j, :));

            if K(j, i) > 1
                theta(i) = min(theta(i), (1 - Kmin) / (K(j, i) - Kmin));
            else
                theta(i) = min(theta(i), (Kmax - 1) / (Kmax - K(j, i)));
            end
        end
    end

    b = min(1 - z./theta, min(1 - K .* z))';

    % Main Newton iteration
    for iter = 1:maxIter
        point(iter, :) = beta';

        % Compute denominator term for each component
        t = (1 - a' * beta)';

        % Compute alpha values
        alpha = a ./ t;

        % Gradient
        grad = alpha * z';

        % Record Residual
        gradNorm = norm(grad, Inf);
        residual(iter) = gradNorm;
        fprintf('Iter %d: Residual = %.2e\n', iter, gradNorm);

        % Check convergence
        if gradNorm < tol
            fprintf('Converged in %d iterations.\n', iter);
            residual = residual(1:iter);
            point = point(1:iter, :);
            return;
        end

        % Hessian matrix
        Hess = alpha * (alpha' .* z');

        % Newton direction
        d = -Hess \ grad;

        % Compute maximum step size allowed by the feasible window
        lambdaMax = 1.0;
        denom = a' * d;
        numer = b - a' * beta;

        for i = 1:NC
            if denom(i) > 0
                lambda = numer(i) / denom(i);
                lambdaMax = max(0, min(lambdaMax, lambda));
            end
        end

        % Line search for step scaling parameter s in [0, 1]
        s = 1.0;
        betaNew = beta;

        for n = 1:maxLineSearchIter
            betaTrial = beta + s * lambdaMax * d;
            tTrial = (1 - a' * betaTrial)';

            alphaTrial = a ./ tTrial;
            gradTrial = alphaTrial * z';

            % First directional derivative
            dg = lambdaMax * (gradTrial' * d);

            if dg < tolLineSearch
                betaNew = betaTrial;
                break;
            end

            % Second directional derivative
            HessTrial = alphaTrial * (alphaTrial' .* z');
            ddg = lambdaMax^2 * (d' * HessTrial * d);

            % Newton update for the line-search parameter
            sNew = s - dg / ddg;
            s = max(0, min(1, sNew));

            betaNew = betaTrial;
        end

        if n == maxLineSearchIter
            warning('Line search did not fully converge. Using s = %.6f.', s);
        end

        % Update beta
        beta = betaNew;
    end

    % Trim unused entries if not converged
    residual = residual(1:maxIter);
    point = point(1:maxIter, :);

    warning('Did not converge after %d iterations.', maxIter);
end