function [beta, iter, residual, point] = RR_Okuno(z, K, beta0, tol, maxIter)
%RACHFORD_RICE_OKUNO Solve the multicomponent Rachford-Rice equations.
%
% Reference:
%   Okuno, Johns, and Sepehrnoori, SPE Journal, 2010.
%
% Inputs:
%   z       : 1 x NC vector of overall compositions
%   K       : (NP-1) x NC matrix of K-values
%   beta0   : (NP-1) x 1 initial guess for phase fractions
%   tol     : convergence tolerance for the gradient norm
%   maxIter : maximum number of Newton iterations
%
% Outputs:
%   beta     : (NP-1) x 1 vector of phase fractions
%   iter     : number of iterations performed
%   residual : history of infinity norms of the gradient
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

    % Storage for convergence history
    residual = zeros(1, maxIter);
    point = zeros(maxIter, length(beta0));

    % Initial guess
    beta = beta0;

    % Precompute coefficients used in the denominator:
    a = 1 - K;
    b = min(1 - z, min(1 - K .* z))';

    % Main Newton iteration loop
    for iter = 1:maxIter
        point(iter, :) = beta';

        % Compute denominator term t_i for each component
        t = (1 - a' * beta)';

        % Compute alpha matrix
        alpha = a ./ t;

        % Gradient of the objective function
        grad = alpha * z';

        % Record residual
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

        % Newton search direction
        d = -Hess \ grad;

        % Compute the maximum allowable step size to stay in the feasible region
        lambdaMax = 1.0;
        denom = a' * d;
        numer = b - a' * beta;

        for i = 1:NC
            if denom(i) > 0
                lambda = numer(i) / denom(i);
                lambdaMax = max(0, min(lambdaMax, lambda));
            end
        end

        % Line search on scalar parameter s in [0,1]
        s = 1.0;
        betaNew = beta;

        for n = 1:maxLineSearchIter
            betaTrial = beta + s * lambdaMax * d;
            tTrial = (1 - a' * betaTrial)';

            alphaTrial = a ./ tTrial;
            gradTrial = alphaTrial * z';

            % First derivative along the search direction
            dg = lambdaMax * (gradTrial' * d);

            if dg < tolLineSearch
                betaNew = betaTrial;
                break;
            end

            % Second derivative along the search direction
            HessTrial = alphaTrial * (alphaTrial' .* z');
            ddg = lambdaMax^2 * (d' * HessTrial * d);

            % Newton update for line search parameter s
            sNew = s - dg / ddg;
            s = max(0, min(1, sNew));

            betaNew = betaTrial;
        end

        if n == maxLineSearchIter
            warning('Line search did not fully converge. Using s = %.6f.', s);
        end

        % Update solution
        beta = betaNew;
    end

    % Trim unused history entries
    residual = residual(1:maxIter);
    point = point(1:maxIter, :);

    warning('Did not converge after %d iterations.', maxIter);
end